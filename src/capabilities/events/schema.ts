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
        "Natural-language search keywords from the user question (English or Burmese). Include a person name when relevant. Sheet/column names vary by workbook — do not assume fixed sheet titles. Leave empty for a light overview.",
    },
    max_results: {
      type: "number",
      description: "Maximum rows to return from hybrid RAG (default 20, keep small).",
    },
  },
};
