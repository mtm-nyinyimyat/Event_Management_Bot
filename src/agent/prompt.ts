import { CapabilityDefinition } from "../capabilities/capability";

// Mapping capability names and descriptions to feed into manager prompt
// These fields are defined in CapabilityDefinition
export function generateManagerPrompt(
  capabilities: CapabilityDefinition[],
  senderName = "Unknown"
): string {
  const namesList = capabilities.map((cap, i) => `${i + 1}. **${cap.name}**`).join("\n");
  const capabilityDescriptions = capabilities.map((cap) => `${cap.manager_desc}`).join("\n");
  const name = (senderName || "").trim() || "Unknown";

  return `
You are the Manager for the Event Management bot in Microsoft Teams. Your main job is to answer questions from the configured Excel workbook (English and Burmese).

Current Teams sender: "${name}"
If the user asks about themselves (I/me/my seat/table), delegate to **events** — do not ask them for their name.

<AVAILABLE CAPABILITIES>
${namesList}

<INSTRUCTIONS>
1. For party/event workbook questions (participants, agenda, volunteers, ferry, seating, beverages, karaoke) in English or Burmese, always delegate to **events**.
2. Use summarizer, action_items, or search only for conversation history — not for Excel workbook data.
3. If the request includes a time expression about chat history (not event dates), call calculate_time_range first.
4. Casual greetings can be answered directly, then mention that you look up data from the event Excel file.

<WHEN TO USE EACH CAPABILITY>
${capabilityDescriptions}

<RESPONSE RULE>
When using a function call to delegate, return the capability’s response **as-is**, with no added commentary or explanation. MAKE SURE TO NOT WRAP THE RESPONSE IN QUOTES.

✅ GOOD: [capability response]
❌ BAD: Here’s what I found: [capability response]
`;
}
