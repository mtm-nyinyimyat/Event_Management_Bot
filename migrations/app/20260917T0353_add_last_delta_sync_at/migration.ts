#!/usr/bin/env -S node
import type { Contract as Start } from '../../snapshots/45171294cfb86c7e16d6205724b8f2cec224bc5cdceb3feefad1eb4178cb9ba7/contract';
import startContract from '../../snapshots/45171294cfb86c7e16d6205724b8f2cec224bc5cdceb3feefad1eb4178cb9ba7/contract.json' with { type: 'json' };
import type { Contract as End } from '../../snapshots/7ff1d3cf141e9dd0ed8a1f2c0880cca598ee9a6edfc0907442f2138ef224a7eb/contract';
import endContract from '../../snapshots/7ff1d3cf141e9dd0ed8a1f2c0880cca598ee9a6edfc0907442f2138ef224a7eb/contract.json' with { type: 'json' };
import { Migration, MigrationCLI, col } from '@prisma/orm-postgres/migration';

export default class M extends Migration<Start, End> {
  override readonly startContractJson = startContract;
  override readonly endContractJson = endContract;

  override get operations() {
    return [
      this.addColumn({
        schema: 'public',
        table: 'rag_sharepoint_sources',
        column: col('last_delta_sync_at', 'timestamptz', {
          codecRef: { codecId: 'pg/timestamptz-string@1' },
        }),
      }),
    ];
  }
}

MigrationCLI.run(import.meta.url, M);
