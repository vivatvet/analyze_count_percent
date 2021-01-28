#!/usr/bin/env python3

from process_acp import Acp
from typing import Tuple


class Catalog(Acp):

    def get_catalog(self, filepath: str) -> Tuple[list, list]:
        return self.load_table(filepath=filepath)

    def rebuild_tabl(self, entry_npp: int, selected_level: str, levels: list, catalog_content: list, raw_tabl: list) -> Tuple[int, list]:
        rebuilded_tabl = []
        i = levels.index(selected_level)
        for rows in raw_tabl:
            for levels_rows in catalog_content:
                if rows[entry_npp] == levels_rows[0]:
                    new_rows = []
                    for item in rows:
                        new_rows.append(item)
                    new_rows.append(levels_rows[i])
                    rebuilded_tabl.append(new_rows)
                    break
        return len(rebuilded_tabl[0]) - 1, rebuilded_tabl
