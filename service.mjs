// @ts-check
import node from "@prisma/composer/node";
import { compute } from "@prisma/composer-prisma-cloud";
import { postgres, dataContract } from "@prisma/composer-prisma-cloud/orm";
import eventmanagementbotContractJson from "./src/prisma/contract.json" with { type: "json" };

export default compute({
  name: "eventmanagementbot",
  deps: { db: postgres(dataContract(eventmanagementbotContractJson)) },
  build: node({ module: import.meta.url, dir: "src", entry: "index.ts" }),
});
