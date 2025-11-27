// index.js (HTTP trigger)
const fetch = require('node-fetch');

const TENANT_ID = process.env.AZ_TENANT_ID;
const CLIENT_ID = process.env.AZ_CLIENT_ID;
const CLIENT_SECRET = process.env.AZ_CLIENT_SECRET;
const ORGANIZER_UPN = process.env.ORGANIZER_UPN; // e.g. teams-bot@contoso.com
const FROM_EMAIL = process.env.FROM_EMAIL || ORGANIZER_UPN;

const tokenEndpoint = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const graphBase = 'https://graph.microsoft.com/v1.0';

async function getAppToken() {
  const body = new URLSearchParams();
  body.append('client_id', CLIENT_ID);
  body.append('scope', 'https://graph.microsoft.com/.default');
  body.append('client_secret', CLIENT_SECRET);
  body.append('grant_type', 'client_credentials');

  const r = await fetch(tokenEndpoint, { method: 'POST', body });
  if (!r.ok) {
    const text = await r.text();
    throw new Error(`Token request failed: ${r.status} ${text}`);
  }
  const json = await r.json();
  return json.access_token;
}

async function createOnlineMeeting(token, { subject, startDateTime, endDateTime }) {
  const url = `${graphBase}/users/${encodeURIComponent(ORGANIZER_UPN)}/onlineMeetings`;
  const body = {
    subject: subject || 'Automated Teams Meeting',
    startDateTime,
    endDateTime
  };

  const r = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(body)
  });

  if (!r.ok) {
    const t = await r.text();
    throw new Error(`Create meeting failed: ${r.status} ${t}`);
  }
  return r.json();
}

async function sendMail(token, { to, subject, bodyHtml }) {
  const url = `${graphBase}/users/${encodeURIComponent(ORGANIZER_UPN)}/sendMail`;

  const mail = {
    message: {
      subject,
      body: {
        contentType: 'HTML',
        content: bodyHtml
      },
      toRecipients: to.map(email => ({ emailAddress: { address: email } }))
    },
    saveToSentItems: "true"
  };

  const r = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(mail)
  });

  if (!r.ok) {
    const t = await r.text();
    throw new Error(`SendMail failed: ${r.status} ${t}`);
  }
  return true;
}

module.exports = async function (context, req) {
  context.log('AutoTeamsCall function invoked');

  try {
    const body = req.body || {};
    const priority = body.priority || 'P1';
    const subject = body.subject || `P1 Incident - ${new Date().toISOString()}`;
    const participants = Array.isArray(body.participants) ? body.participants : [];
    const startInMinutes = Number(body.startInMinutes || 1);
    const durationMinutes = Number(body.durationMinutes || 30);

    if (participants.length === 0) {
      context.log('No participants provided');
      return context.res = { status: 400, body: 'participants array required' };
    }

    const token = await getAppToken();

    const start = new Date(Date.now() + Math.max(0, startInMinutes) * 60 * 1000);
    const end = new Date(start.getTime() + Math.max(1, durationMinutes) * 60 * 1000);

    const meeting = await createOnlineMeeting(token, {
      subject,
      startDateTime: start.toISOString(),
      endDateTime: end.toISOString()
    });

    const joinUrl =
      meeting.joinUrl ||
      (meeting.joinInformation && meeting.joinInformation.joinUrl) ||
      null;

    if (!joinUrl) {
      throw new Error('No joinUrl returned from Graph meeting creation');
    }

    const html = `
      <p>Priority: <b>${priority}</b></p>
      <p>${subject}</p>
      <p>Start: ${start.toISOString()}</p>
      <p>Join the Teams meeting: <a href="${joinUrl}">Join meeting</a></p>
      <p>Link: ${joinUrl}</p>
    `;

    await sendMail(token, {
      to: participants,
      subject: `[${priority}] ${subject} — Teams meeting`,
      bodyHtml: html
    });

    return context.res = {
      status: 200,
      body: {
        meetingId: meeting.id,
        joinUrl
      }
    };

  } catch (err) {
    context.log.error('Function error:', err);
    return context.res = {
      status: 500,
      body: { error: err.message || String(err) }
    };
  }
};

// // TeamsCall/index.js - minimal test handler
// module.exports = async function (context, req) {
//   context.log('TeamsCall invoked (minimal handler)');
//   context.res = {
//     status: 200,
//     body: {
//       ok: true,
//       message: 'TeamsCall function reached (local test)',
//       received: req.body || null
//     }
//   };
// };
