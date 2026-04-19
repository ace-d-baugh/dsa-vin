const { BotFrameworkAdapter, TurnContext } = require('botbuilder');
const restify = require('restify');
const fs = require('fs').promises;
const path = require('path');
require('dotenv').config();

// ─── Bot Adapter ─────────────────────────────────────────────────────────────

const adapter = new BotFrameworkAdapter({
  appId: process.env.MicrosoftAppId,
  appPassword: process.env.MicrosoftAppPassword,
});

adapter.onTurnError = async (context, error) => {
  console.error('Bot error:', error.message);
};

// ─── HTTP Server ──────────────────────────────────────────────────────────────

const server = restify.createServer();
server.use(restify.plugins.bodyParser());

server.post('/api/messages', (req, res) => {
  adapter.processActivity(req, res, async (context) => {
    await handleTurn(context);
  });
});

const PORT = process.env.PORT || 3978;
server.listen(PORT, () => {
  console.log(`DSA VIN Teams bot listening on port ${PORT}`);
});

// ─── Config ───────────────────────────────────────────────────────────────────

const CONFIG = {
  API_URL: 'https://api.themeparks.wiki/v1/entity/e957da41-3552-4cf6-b636-5babc5cbc4e5/live',
  POLL_INTERVAL: 10000,
  CONV_REF_FILE: path.join(__dirname, 'conversationRefs.json'),
  STATE_FILE: path.join(__dirname, 'attractionState.json'),
};

// WDW attractions — all four parks
const ATTRACTIONS = [
  // Magic Kingdom
  { name: 'Astro Orbiter', id: 'd9d12438-d999-4482-894b-8955fdb20ccf', park: 'MK' },
  { name: 'The Barnstormer', id: '924a3b2c-6b4b-49e5-99d3-e9dc3f2e8a48', park: 'MK' },
  { name: 'Big Thunder Mountain Railroad', id: 'de3309ca-97d5-4211-bffe-739fed47e92f', park: 'MK' },
  { name: "Buzz Lightyear's Space Ranger Spin", id: '72c7343a-f7fb-4f66-95df-c91016de7338', park: 'MK' },
  { name: 'The Magic Carpets of Aladdin', id: '96455de6-f4f1-403c-9391-bf8396979149', park: 'MK' },
  { name: "Walt Disney's Carousel of Progress", id: '8183f3f2-1b59-4b9c-b634-6a863bdf8d84', park: 'MK' },
  { name: 'Country Bear Musical Jamboree', id: '0f57cecf-5502-4503-8bc3-ba84d3708ace', park: 'MK' },
  { name: 'Dumbo the Flying Elephant', id: '890fa430-89c0-4a3f-96c9-11597888005e', park: 'MK' },
  { name: 'Enchanted Tales with Belle', id: 'e76c93df-31af-49a5-8e2f-752c76c937c9', park: 'MK' },
  { name: 'The Hall of Presidents', id: '2ebfb38c-5cb5-4de1-86c0-f7af14188022', park: 'MK' },
  { name: 'Haunted Mansion', id: '2551a77d-023f-4ab1-9a19-8afec0190f39', park: 'MK' },
  { name: 'Jungle Cruise', id: '796b0a25-c51e-456e-9bb8-50a324e301b3', park: 'MK' },
  { name: 'Monsters Inc. Laugh Floor', id: 'e8f0b426-7645-4ea3-8b41-b94ae7091a41', park: 'MK' },
  { name: 'Under the Sea – Journey of The Little Mermaid', id: '3cba0cb4-e2a6-402c-93ee-c11ffcb127ef', park: 'MK' },
  { name: 'Seven Dwarfs Mine Train', id: '9d4d5229-7142-44b6-b4fb-528920969a2c', park: 'MK' },
  { name: "Peter Pan's Flight", id: '86a41273-5f15-4b54-93b6-829f140e5161', park: 'MK' },
  { name: 'Tomorrowland Transit Authority PeopleMover', id: 'ffcfeaa2-1416-4920-a1ed-543c1a1695c4', park: 'MK' },
  { name: "Mickey's PhilharMagic", id: '7c5e1e02-3a44-4151-9005-44066d5ba1da', park: 'MK' },
  { name: 'Pirates of the Caribbean', id: '352feb94-e52e-45eb-9c92-e4b44c6b1a9d', park: 'MK' },
  { name: 'The Many Adventures of Winnie the Pooh', id: '0d94ad60-72f0-4551-83a6-ebaecdd89737', park: 'MK' },
  { name: "it's a small world", id: 'f5aad2d4-a419-4384-bd9a-42f86385c750', park: 'MK' },
  { name: 'Space Mountain', id: 'b2260923-9315-40fd-9c6b-44dd811dbe64', park: 'MK' },
  { name: 'Tomorrowland Speedway', id: 'f163ddcd-43e1-488d-8276-2381c1db0a39', park: 'MK' },
  { name: 'Mad Tea Party', id: '0aae716c-af13-4439-b638-d75fb1649df3', park: 'MK' },
  { name: "Tiana's Bayou Adventure", id: '73cb9445-0695-47a3-87ce-d08ae36b5f3c', park: 'MK' },
  { name: "Walt Disney's Enchanted Tiki Room", id: '6fd1e225-53a0-4a80-a577-4bbc9a471075', park: 'MK' },
  { name: 'TRON Lightcycle / Run', id: '5a43d1a7-ad53-4d25-abfe-25625f0da304', park: 'MK' },

  // Epcot
  { name: 'The American Adventure', id: '1f542745-cda1-4786-a536-5fff373e5964', park: 'EP' },
  { name: 'Journey Into Imagination With Figment', id: '75449e85-c410-4cef-a368-9d2ea5d52b58', park: 'EP' },
  { name: 'Frozen Ever After', id: '8d7ccdb1-a22b-4e26-8dc8-65b1938ed5f0', park: 'EP' },
  { name: 'Gran Fiesta Tour', id: '22f48b73-01df-460e-8969-9eb2b4ae836c', park: 'EP' },
  { name: 'Guardians of the Galaxy: Cosmic Rewind', id: 'e3549451-b284-453d-9c31-e3b1207abd79', park: 'EP' },
  { name: 'Living with the Land', id: '8f353879-d6ac-4211-9352-4029efb47c18', park: 'EP' },
  { name: 'Mission: SPACE', id: '5b6475ad-4e9a-4793-b841-501aa382c9c0', park: 'EP' },
  { name: 'Journey of Water, Inspired by Moana', id: 'dae68dee-dfba-4128-b594-6aa12add1070', park: 'EP' },
  { name: 'The Seas with Nemo & Friends', id: 'fb076275-0570-4d62-b2a9-4d6515130fa3', park: 'EP' },
  { name: "Remy's Ratatouille Adventure", id: '1e735ffb-4868-47f1-b2cd-2ac1156cd5f0', park: 'EP' },
  { name: "Soarin' Around the World", id: '81b15dfd-cf6a-466f-be59-3dd65d2a2807', park: 'EP' },
  { name: 'Spaceship Earth', id: '480fde8f-fe58-4bfb-b3ab-052a39d4db7c', park: 'EP' },
  { name: 'Test Track', id: '37ae57c5-feaf-4e47-8f27-4b385be200f0', park: 'EP' },
  { name: 'Turtle Talk With Crush', id: '57acb522-a6fc-4aa4-a80e-21f21f317250', park: 'EP' },

  // Hollywood Studios
  { name: "Rock 'n' Roller Coaster Starring Aerosmith", id: 'e516f303-e82d-4fd3-8fbf-8e6ab624cf89', park: 'HS' },
  { name: 'Toy Story Mania!', id: '20b5daa8-e1ea-436f-830c-2d7d18d929b5', park: 'HS' },
  { name: "Mickey & Minnie's Runaway Railway", id: '6e118e37-5002-408d-9d88-0b5d9cdb5d14', park: 'HS' },
  { name: 'Star Wars: Rise of the Resistance', id: '1a2e70d9-50d5-4140-b69e-799e950f7d18', park: 'HS' },
  { name: 'Alien Swirling Saucers', id: 'd56506e2-6ad3-443a-8065-fea37987248d', park: 'HS' },
  { name: 'Slinky Dog Dash', id: '399aa0a1-98e2-4d2b-b297-2b451e9665e1', park: 'HS' },
  { name: 'Millennium Falcon: Smugglers Run', id: '34c4916b-989b-4ff1-a7e3-a6a846a3484f', park: 'HS' },
  { name: 'Star Tours – The Adventures Continue', id: '3b290419-8ca2-44bc-a710-a6c83fca76ec', park: 'HS' },
  { name: 'The Twilight Zone Tower of Terror', id: '6f6998e8-a629-412c-b964-2cb06af8e26b', park: 'HS' },

  // Animal Kingdom
  { name: 'Expedition Everest', id: '64a6915f-a835-4226-ba5c-8389fc4cade3', park: 'AK' },
  { name: 'Avatar Flight of Passage', id: '24cf863c-b6ba-4826-a056-0b698989cbf7', park: 'AK' },
  { name: 'Kali River Rapids', id: 'd58d9262-ec95-4161-80a0-07ca43b2f5f3', park: 'AK' },
  { name: 'Kilimanjaro Safaris', id: '32e01181-9a5f-4936-8a77-0dace1de836c', park: 'AK' },
  { name: "Na'vi River Journey", id: '7a5af3b7-9bc1-4962-92d0-3ea9c9ce35f0', park: 'AK' },
  { name: 'Wildlife Express Train', id: '4f391f0e-52be-4f9d-99d6-b3ae0373b43c', park: 'AK' },
  { name: 'Zootopia: Better Zoogether!', id: '1b15c77b-0311-4171-8e59-7f38e6d60754', park: 'AK' },
];

const PARK_NAMES = { MK: 'Magic Kingdom', EP: 'Epcot', HS: "Hollywood Studios", AK: 'Animal Kingdom' };

// ─── Conversation Reference Store ─────────────────────────────────────────────
// Persists all chats/channels the bot has been added to so proactive messages
// can be sent without an active incoming turn.

async function loadConversationRefs() {
  try {
    const raw = await fs.readFile(CONFIG.CONV_REF_FILE, 'utf8');
    return JSON.parse(raw);
  } catch {
    return {};
  }
}

async function saveConversationRefs(refs) {
  await fs.writeFile(CONFIG.CONV_REF_FILE, JSON.stringify(refs, null, 2));
}

async function storeConversationRef(context) {
  const ref = TurnContext.getConversationReference(context.activity);
  const key = ref.conversation.id;
  const refs = await loadConversationRefs();
  if (!refs[key]) {
    refs[key] = ref;
    await saveConversationRefs(refs);
    console.log(`Stored conversation reference: ${key}`);
  }
}

// ─── Attraction State ─────────────────────────────────────────────────────────

async function loadState() {
  try {
    const raw = await fs.readFile(CONFIG.STATE_FILE, 'utf8');
    return JSON.parse(raw);
  } catch {
    // Build initial state — all unknown so first real poll doesn't false-alarm
    const state = {};
    for (const attr of ATTRACTIONS) {
      state[attr.id] = { status: null };
    }
    return state;
  }
}

async function saveState(state) {
  await fs.writeFile(CONFIG.STATE_FILE, JSON.stringify(state, null, 2));
}

// ─── Bot Turn Handler ─────────────────────────────────────────────────────────

async function handleTurn(context) {
  const type = context.activity.type;

  if (type === 'installationUpdate' || type === 'conversationUpdate') {
    // Bot was added to a chat — store the reference and greet
    await storeConversationRef(context);
    if (
      type === 'installationUpdate' ||
      (context.activity.membersAdded || []).some(m => m.id === context.activity.recipient.id)
    ) {
      await context.sendActivity(
        '👋 **DSA VIN** is online! I\'ll post here whenever a Walt Disney World attraction goes down (❌ 101) or comes back up (✅ 102).'
      );
    }
  } else if (type === 'message') {
    // Store ref on any message too (covers group chats where installationUpdate may not fire)
    await storeConversationRef(context);
    const text = (context.activity.text || '').trim().toLowerCase();
    if (text === 'status' || text === 'dsa vin status') {
      await context.sendActivity('✅ DSA VIN is running and polling every 10 seconds.');
    }
  }
}

// ─── Proactive Message Broadcast ──────────────────────────────────────────────

async function broadcast(message) {
  const refs = await loadConversationRefs();
  for (const ref of Object.values(refs)) {
    try {
      await adapter.continueConversation(ref, async (context) => {
        await context.sendActivity(message);
      });
    } catch (err) {
      console.error(`Broadcast failed for conversation ${ref.conversation.id}:`, err.message);
    }
  }
}

// ─── Polling Loop ─────────────────────────────────────────────────────────────

function formatTime() {
  return new Date().toLocaleTimeString('en-US', {
    hour: 'numeric',
    minute: '2-digit',
    hour12: true,
    timeZone: 'America/New_York',
  });
}

let isStartup = true;
let isPolling = false;

async function poll() {
  if (isPolling) return;
  isPolling = true;

  try {
    const response = await fetch(CONFIG.API_URL);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const data = await response.json();

    const state = await loadState();
    const time = formatTime();
    const changed = [];

    for (const attr of ATTRACTIONS) {
      const live = data.liveData.find(item => item.id === attr.id);
      if (!live) continue;

      const prev = state[attr.id]?.status ?? null;
      const curr = live.status; // 'OPERATING', 'DOWN', 'CLOSED', 'REFURBISHMENT'

      if (curr === prev) continue; // No change

      state[attr.id] = { status: curr };

      // Only notify on DOWN and recovery-to-OPERATING transitions
      if (!isStartup) {
        const parkLabel = PARK_NAMES[attr.park] || attr.park;

        if (curr === 'DOWN' && prev === 'OPERATING') {
          changed.push(`❌ **101 — DOWN** | ${attr.name} | ${parkLabel} | ${time}`);
        } else if (curr === 'OPERATING' && prev === 'DOWN') {
          changed.push(`✅ **102 — UP** | ${attr.name} | ${parkLabel} | ${time}`);
        }
        // CLOSED and REFURBISHMENT changes are silent — add cases here if you want them
      }
    }

    await saveState(state);

    for (const msg of changed) {
      console.log(msg);
      await broadcast(msg);
    }

    if (isStartup) {
      console.log(`Startup sync complete — tracking ${ATTRACTIONS.length} attractions.`);
      isStartup = false;
    }
  } catch (err) {
    console.error('Poll error:', err.message);
  } finally {
    isPolling = false;
  }
}

// Start polling after a short delay to let the server come up
setTimeout(() => {
  poll();
  setInterval(poll, CONFIG.POLL_INTERVAL);
}, 3000);
