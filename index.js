const { BotFrameworkAdapter, ActivityHandler, MessageFactory } = require('botbuilder');
const restify = require('restify');
require('dotenv').config();

// ── Adapter ──────────────────────────────────────────────────────────────────
const adapter = new BotFrameworkAdapter({
  appId: process.env.MICROSOFT_APP_ID || '',
  appPassword: process.env.MICROSOFT_APP_PASSWORD || '',
});

adapter.onTurnError = async (context, error) => {
  console.error('[BreakBot Error]', error);
  await context.sendActivity('Something went wrong. Please try again.');
};

// ── In-memory store (persists while server is running) ───────────────────────
// Structure: { [memberName]: { onBreak: bool, start: Date, totalMins: number, breakCount: number } }
const memberState = {};
// Full log for timesheet
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

// ── Bot Logic ─────────────────────────────────────────────────────────────────
class BreakBot extends ActivityHandler {
  constructor() {
    super();

    this.onMembersAdded(async (context, next) => {
      for (const member of context.activity.membersAdded) {
        if (member.id !== context.activity.recipient.id) {
          await context.sendActivity(
            `👋 **BreakBot is active!**\n\n` +
            `I track your team's break times automatically.\n\n` +
            `**Commands:**\n` +
            `• \`Break\` — start your break\n` +
            `• \`Back from Break\` — end your break\n` +
            `• \`Status\` — check your current status\n` +
            `• \`Report\` — see today's full team timesheet\n` +
            `• \`My Breaks\` — see your personal break history`
          );
        }
      }
      await next();
    });

    this.onMessage(async (context, next) => {
      const member = context.activity.from.name || 'Unknown';
      const text = (context.activity.text || '').trim().toLowerCase();

      // ── BREAK ──────────────────────────────────────────────────────────────
      if (text === 'break') {
        const s = memberState[member];
        if (s && s.onBreak) {
          await context.sendActivity(
            `⚠️ **${member}**, you're already on a break (started at ${s.startStr}).\n` +
            `Send \`Back from Break\` to end it.`
          );
        } else {
          memberState[member] = {
            onBreak: true,
            start: Date.now(),
            startStr: nowStr(),
            totalMins: (s && s.totalMins) || 0,
            breakCount: ((s && s.breakCount) || 0) + 1,
          };
          breakLog.push({
            member,
            date: nowDate(),
            breakStart: nowStr(),
            breakEnd: null,
            duration: null,
          });
          await context.sendActivity(
            `☕ **${member}** is on a break!\n` +
            `_Started at ${nowStr()}_\n\n` +
            `Send \`Back from Break\` when you return.`
          );
        }

      // ── BACK FROM BREAK ────────────────────────────────────────────────────
      } else if (text === 'back from break' || text === 'back') {
        const s = memberState[member];
        if (!s || !s.onBreak) {
          await context.sendActivity(
            `ℹ️ **${member}**, you don't have an active break.\n` +
            `Send \`Break\` to start one.`
          );
        } else {
          const mins = elapsed(s.start);
          const endStr = nowStr();
          memberState[member] = {
            onBreak: false,
            totalMins: s.totalMins + mins,
            breakCount: s.breakCount,
          };
          // Update last log entry for this member
          const entry = [...breakLog].reverse().find(e => e.member === member && !e.breakEnd);
          if (entry) { entry.breakEnd = endStr; entry.duration = mins; }

          await context.sendActivity(
            `✅ Welcome back, **${member}**!\n` +
            `_Break ended at ${endStr}_ — Duration: **${mins} min**\n\n` +
            `Total break time today: **${memberState[member].totalMins} min**`
          );
        }

      // ── STATUS ─────────────────────────────────────────────────────────────
      } else if (text === 'status') {
        const s = memberState[member];
        if (!s) {
          await context.sendActivity(`📋 **${member}** — no activity logged today.`);
        } else if (s.onBreak) {
          const mins = elapsed(s.start);
          await context.sendActivity(
            `🟡 **${member}** is currently **on break**\n` +
            `Started: ${s.startStr} | Elapsed: **${mins} min**`
          );
        } else {
          await context.sendActivity(
            `🟢 **${member}** is **active**\n` +
            `Total break time today: **${s.totalMins} min** | Breaks taken: **${s.breakCount}**`
          );
        }

      // ── MY BREAKS ──────────────────────────────────────────────────────────
      } else if (text === 'my breaks') {
        const myLogs = breakLog.filter(e => e.member === member);
        if (!myLogs.length) {
          await context.sendActivity(`📋 **${member}** — no break history found.`);
        } else {
          let msg = `📋 **${member}'s Break History (Today)**\n\n`;
          myLogs.forEach((e, i) => {
            msg += `${i + 1}. Start: ${e.breakStart} | End: ${e.breakEnd || '⏳ ongoing'} | Duration: ${e.duration ? e.duration + ' min' : '—'}\n`;
          });
          await context.sendActivity(msg);
        }

      // ── REPORT ─────────────────────────────────────────────────────────────
      } else if (text === 'report') {
        if (!Object.keys(memberState).length) {
          await context.sendActivity(`📊 No break data recorded yet today.`);
        } else {
          let msg = `📊 **Team Break Report — ${nowDate()}**\n\n`;
          for (const [name, s] of Object.entries(memberState)) {
            const status = s.onBreak ? '🟡 On Break' : '🟢 Active';
            const total = s.onBreak ? (s.totalMins + elapsed(s.start)) : s.totalMins;
            msg += `**${name}** — ${status} | Total: **${total} min** | Breaks: **${s.breakCount}**\n`;
          }
          await context.sendActivity(msg);
        }

      // ── HELP / UNKNOWN ─────────────────────────────────────────────────────
      } else if (text === 'help') {
        await context.sendActivity(
          `**BreakBot Commands:**\n\n` +
          `• \`Break\` — start your break\n` +
          `• \`Back from Break\` — end your break\n` +
          `• \`Status\` — your current status\n` +
          `• \`My Breaks\` — your break history today\n` +
          `• \`Report\` — full team timesheet`
        );
      } else {
        await context.sendActivity(
          `❓ I didn't understand that.\n\n` +
          `Try: **Break**, **Back from Break**, **Status**, **Report**, or **Help**`
        );
      }

      await next();
    });
  }
}

// ── Server ────────────────────────────────────────────────────────────────────
const bot = new BreakBot();
const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.post('/api/messages', (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

// Health check endpoint
server.get('/health', (req, res, next) => {
  res.send(200, { status: 'BreakBot running', time: new Date().toISOString() });
  return next();
});

const PORT = process.env.PORT || 3978;
server.listen(PORT, () => {
  console.log(`✅ BreakBot running on port ${PORT}`);
  console.log(`   Messaging endpoint: /api/messages`);
});
