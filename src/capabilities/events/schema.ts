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
        "Keywords in English or Burmese to match any sheet/column (member name, dance group, date, location, costume, cost). Leave empty for an overview.",
    },
    max_results: {
      type: "number",
      description: "Maximum matching rows to return. Defaults to 20.",
    },
  },
  required: [],
};
