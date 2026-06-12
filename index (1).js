const http = require('http');
const crypto = require('crypto');

const PORT = process.env.PORT || 3978;
const APP_ID = process.env.MICROSOFT_APP_ID || '';
const APP_PASSWORD = process.env.MICROSOFT_APP_PASSWORD || '';

// ── In-memory store ──────────────────────────────────────────────────────────
const memberState = {};
const breakLog = [];

function nowStr() {
  return new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
}
function nowDate() {
  return new Date().toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
}
function elapsed(start) {
  return Math.round((Date.now() - start) / 60000);
}

// ── Get Bot Framework token ───────────────────────────────────────────────────
async function getToken() {
  const body = `grant_type=client_credentials&client_id=${encodeURIComponent(APP_ID)}&client_secret=${encodeURIComponent(APP_PASSWORD)}&scope=https%3A%2F%2Fapi.botframework.com%2F.default`;
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'login.microsoftonline.com',
      path: '/botframework.com/oauth2/v2.0/token',
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) }
    }, res => {
      let data = '';
      res.on('data', d => data += d);
      res.on('end', () => {
        try { resolve(JSON.parse(data).access_token); } catch { reject(new Error('Token parse failed')); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

const https = require('https');

// ── Send reply back to Teams ──────────────────────────────────────────────────
async function sendReply(activity, text) {
  try {
    const token = await getToken();
    const serviceUrl = activity.serviceUrl.replace(/\/$/, '');
    const conversationId = activity.conversation.id;
    const replyBody = JSON.stringify({
      type: 'message',
      from: { id: activity.recipient.id, name: activity.recipient.name },
      conversation: activity.conversation,
      recipient: activity.from,
      text: text,
      replyToId: activity.id
    });

    const url = new URL(`${serviceUrl}/v3/conversations/${encodeURIComponent(conversationId)}/activities`);
    return new Promise((resolve, reject) => {
      const req = https.request({
        hostname: url.hostname,
        path: url.pathname + url.search,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`,
          'Content-Length': Buffer.byteLength(replyBody)
        }
      }, res => {
        let d = '';
        res.on('data', c => d += c);
        res.on('end', () => resolve(d));
      });
      req.on('error', reject);
      req.write(replyBody);
      req.end();
    });
  } catch (err) {
    console.error('sendReply error:', err.message);
  }
}

// ── Bot logic ─────────────────────────────────────────────────────────────────
async function handleMessage(activity) {
  const member = (activity.from && activity.from.name) ? activity.from.name : 'Unknown';
  const text = (activity.text || '').trim().toLowerCase().replace(/<[^>]+>/g, '');

  let reply = '';

  if (text === 'break') {
    const s = memberState[member];
    if (s && s.onBreak) {
      reply = `⚠️ ${member}, you're already on a break (started ${s.startStr}). Send "Back from Break" to end it.`;
    } else {
      memberState[member] = { onBreak: true, start: Date.now(), startStr: nowStr(), totalMins: (s && s.totalMins) || 0, breakCount: ((s && s.breakCount) || 0) + 1 };
      breakLog.push({ member, date: nowDate(), breakStart: nowStr(), breakEnd: null, duration: null });
      reply = `☕ ${member} is on a break! Started at ${nowStr()}. Send "Back from Break" when you return.`;
    }

  } else if (text === 'back from break' || text === 'back') {
    const s = memberState[member];
    if (!s || !s.onBreak) {
      reply = `ℹ️ ${member}, you don't have an active break. Send "Break" to start one.`;
    } else {
      const mins = elapsed(s.start);
      const endStr = nowStr();
      memberState[member] = { onBreak: false, totalMins: s.totalMins + mins, breakCount: s.breakCount };
      const entry = [...breakLog].reverse().find(e => e.member === member && !e.breakEnd);
      if (entry) { entry.breakEnd = endStr; entry.duration = mins; }
      reply = `✅ Welcome back, ${member}! Break ended at ${endStr}. Duration: ${mins} min. Total today: ${memberState[member].totalMins} min.`;
    }

  } else if (text === 'status') {
    const s = memberState[member];
    if (!s) {
      reply = `📋 ${member} — no activity logged today.`;
    } else if (s.onBreak) {
      reply = `🟡 ${member} is on break — started ${s.startStr}, ${elapsed(s.start)} min so far.`;
    } else {
      reply = `🟢 ${member} is active. Total break time today: ${s.totalMins} min (${s.breakCount} breaks).`;
    }

  } else if (text === 'report') {
    if (!Object.keys(memberState).length) {
      reply = '📊 No break data recorded yet today.';
    } else {
      let msg = `📊 Team Break Report — ${nowDate()}\n\n`;
      for (const [name, s] of Object.entries(memberState)) {
        const status = s.onBreak ? '🟡 On Break' : '🟢 Active';
        const total = s.onBreak ? (s.totalMins + elapsed(s.start)) : s.totalMins;
        msg += `${name} — ${status} | Total: ${total} min | Breaks: ${s.breakCount}\n`;
      }
      reply = msg;
    }

  } else if (text === 'my breaks') {
    const myLogs = breakLog.filter(e => e.member === member);
    if (!myLogs.length) {
      reply = `📋 ${member} — no break history found.`;
    } else {
      let msg = `📋 ${member}'s Break History (Today)\n\n`;
      myLogs.forEach((e, i) => {
        msg += `${i + 1}. Start: ${e.breakStart} | End: ${e.breakEnd || '⏳ ongoing'} | Duration: ${e.duration ? e.duration + ' min' : '—'}\n`;
      });
      reply = msg;
    }

  } else if (text === 'help' || text === 'hi' || text === 'hello') {
    reply = `👋 Hi ${member}! I'm BreakBot.\n\nCommands:\n• "Break" — start your break\n• "Back from Break" — end your break\n• "Status" — your current status\n• "My Breaks" — your break history\n• "Report" — full team timesheet`;
  } else {
    reply = `❓ I didn't understand that. Try: Break, Back from Break, Status, Report, or Help.`;
  }

  await sendReply(activity, reply);
}

// ── HTTP Server ───────────────────────────────────────────────────────────────
const server = http.createServer((req, res) => {
  if (req.method === 'GET' && req.url === '/health') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ status: 'BreakBot running', time: new Date().toISOString() }));
    return;
  }

  if (req.method === 'POST' && req.url === '/api/messages') {
    let body = '';
    req.on('data', chunk => body += chunk);
    req.on('end', async () => {
      try {
        const activity = JSON.parse(body);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ status: 'ok' }));
        if (activity.type === 'message') {
          await handleMessage(activity);
        } else if (activity.type === 'conversationUpdate') {
          if (activity.membersAdded && activity.membersAdded.some(m => m.id !== activity.recipient.id)) {
            await sendReply(activity, `👋 BreakBot is active! Type "Help" to see all commands.`);
          }
        }
      } catch (err) {
        console.error('Error:', err.message);
        res.writeHead(200);
        res.end('ok');
      }
    });
    return;
  }

  res.writeHead(404);
  res.end('Not found');
});

server.listen(PORT, () => {
  console.log(`✅ BreakBot running on port ${PORT}`);
});
