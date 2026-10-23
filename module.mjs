// @ts-check
import { module } from "@prisma/composer";
import { postgres, dataContract } from "@prisma/composer-prisma-cloud/orm";
import eventmanagementbotContractJson from "./src/prisma/contract.json" with { type: "json" };
import eventmanagementbotService from "./service.mjs";

export default module("event-management-bot", ({ provision }) => {
  const database = provision(postgres({ name: "database", contract: dataContract(eventmanagementbotContractJson), config: "./prisma.config.ts" }));
  provision(eventmanagementbotService, { deps: { db: database } });
});
