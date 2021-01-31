#!/usr/bin/env python3

from process_acp import Acp
from typing import Tuple


class Catalog(Acp):

    def get_catalog(self, filepath: str) -> Tuple[list, list]:
        return self.load_table(filepath=filepath)

    @staticmethod
    def rebuild_tabl(entry_npp: int, selected_level: str, levels: list, catalog_content: list, raw_tabl: list) -> Tuple[bool, int, list, list]:
        rebuilded_tabl = []
        not_found_code = []
        i = levels.index(selected_level)
        for rows in raw_tabl:
            found = False
            for levels_rows in catalog_content:
                if rows[entry_npp] == levels_rows[0]:
                    new_rows = []
                    for item in rows:
                        new_rows.append(item)
                    new_rows.append(levels_rows[i])
                    rebuilded_tabl.append(new_rows)
                    found = True
                    break
            if not found:
                not_found_code.append(rows[entry_npp])

        if 0 < len(not_found_code) <= 20:
            return False, 0, [], not_found_code
        elif len(not_found_code) > 20:
            return False, 0, [], ["more then 20 codes..."]

        return True, len(rebuilded_tabl[0]) - 1, rebuilded_tabl, []

    @staticmethod
    def rebuild_tabl_by_market(entry_rf: int, raw_tabl: list, rf: str) -> list:
        rebuilded_tabl = []
        for row in raw_tabl:
            if row[entry_rf] == rf:
                rebuilded_tabl.append(row)
        return rebuilded_tabl
