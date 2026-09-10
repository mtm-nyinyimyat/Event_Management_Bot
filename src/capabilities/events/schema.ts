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
        "Natural-language keywords in English or Burmese. Prefer a sheet focus word: Agenda, Participants, Volunteer, Ferry, Table. For a person, use their name or a short question. Leave empty for sheet summaries only.",
    },
    max_results: {
      type: "number",
      description:
        "Maximum rows to return per sheet when listing. Count/total questions can keep the default; use up to 80 for full lists. Defaults to 80.",
    },
  },
  required: [],
};
