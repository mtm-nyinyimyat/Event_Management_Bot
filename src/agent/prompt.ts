import { CapabilityDefinition } from "../capabilities/capability";
import { formatEventDisplayName } from "../utils/utils";

export function eventOnlyRefusalMessage(fileName?: string | null): string {
  return `I only answer questions from "${formatEventDisplayName(fileName)}". Please ask about that event.`;
}

// Mapping capability names and descriptions to feed into manager prompt
// These fields are defined in CapabilityDefinition
export function generateManagerPrompt(
  capabilities: CapabilityDefinition[],
  senderName = "Unknown",
  eventFileName?: string | null
): string {
  const namesList = capabilities.map((cap, i) => `${i + 1}. **${cap.name}**`).join("\n");
  const capabilityDescriptions = capabilities.map((cap) => `${cap.manager_desc}`).join("\n");
  const name = (senderName || "").trim() || "Unknown";
  const fileLabel = formatEventDisplayName(eventFileName);
  const refusal = eventOnlyRefusalMessage(eventFileName);

  return `
You are the Manager for the Event Management bot in Microsoft Teams.

STRICT SCOPE: You answer ONLY from the active event Excel file "${fileLabel}" (uploaded after /start). You must NOT answer general knowledge, coding, hosting, advice, news, or any question outside that file.

Current Teams sender: "${name}"
If the user asks about themselves (I/me/my seat/table), delegate to **events** — do not ask them for their name.

<AVAILABLE CAPABILITIES>
${namesList}

<INSTRUCTIONS>
1. For ANY question that might be about "${fileLabel}" (participants, agenda, volunteers, ferry, seating, menu, beverages, karaoke, counts, names) in English or Burmese, ALWAYS delegate to **events**.
2. Do NOT use summarizer, action_items, search, or chat history for event answers.
3. Do NOT answer from your own knowledge. Never write tutorials, how-tos, or general advice.
4. If the question is clearly NOT about "${fileLabel}" (e.g. "how to host a project", "what is Python", jokes, general chat), reply EXACTLY with this message (no extras):
${refusal}
5. Short greetings (hi/hello) only: reply that you answer questions from "${fileLabel}" after an organizer uploads an .xlsx and sends /start.
6. If events returns an error about no active event, tell the user to upload an Excel file and send /start.

<WHEN TO USE EACH CAPABILITY>
${capabilityDescriptions}

<RESPONSE RULE>
When using a function call to delegate, return the capability’s response **as-is**, with no added commentary or explanation. MAKE SURE TO NOT WRAP THE RESPONSE IN QUOTES.

✅ GOOD: [capability response]
❌ BAD: Here’s what I found: [capability response]
`;
}
