import { ObjectSchema } from "@microsoft/teams.ai";

export interface LookupEventsArgs {
  query?: string;
  max_results?: number;
}

export const LOOKUP_EVENTS_SCHEMA: ObjectSchema = {
  type: "object",
  properties: {
    query: {
      type: "string",
      description:
        "Natural-language keywords in English or Burmese. Prefer a sheet focus word: Agenda, Participants, Volunteer, Ferry, Table. For drinks use beer/juice/cocktail/beverage. For a person, use their name or a short question. Leave empty for sheet summaries only.",
    },
    max_results: {
      type: "number",
      description:
        "Maximum rows to return. Keep small for fuzzy RAG (default 20). Beverage/menu list queries return the full filtered set from the tool regardless.",
    },
  },
};
