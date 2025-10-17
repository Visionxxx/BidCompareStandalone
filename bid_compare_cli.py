#!/usr/bin/env python3
"""
Bid Compare CLI - Sammenlign anbudstilbud fra kommandolinjen

Bruk:
    python bid_compare_cli.py tilbud1.xlsx tilbud2.xml tilbud3.csv
    python bid_compare_cli.py --output rapport.xlsx tilbud*.xml
"""

import argparse
import sys
from pathlib import Path

# Import fra backend
sys.path.insert(0, str(Path(__file__).parent / "backend"))
from backend.app.main import (
    _parse_ns3459_xml,
    _read_tabular,
    _normalize_columns,
    _extract_company_name,
    _collect_chapter_titles,
    _aggregate_bid_rows,
    _to_float,
)

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font


def load_bid_file(filepath: Path) -> tuple[str, pd.DataFrame]:
    """Last inn en tilbudsfil og returner navn + DataFrame"""
    data = filepath.read_bytes()

    if filepath.suffix.lower() == '.xml':
        df = _parse_ns3459_xml(data, filepath.name)
        try:
            name = _extract_company_name(data) or filepath.stem
        except Exception:
            name = filepath.stem
    else:
        df_raw = _read_tabular(filepath.name, data)
        df = _normalize_columns(df_raw)
        name = filepath.stem

    return name, df


def print_summary(bids: dict[str, pd.DataFrame], base_bids: dict[str, pd.DataFrame], option_totals: dict[str, float]):
    """Skriv ut oppsummering til terminal"""
    print("\n" + "=" * 80)
    print("OPPSUMMERING")
    print("=" * 80)

    totals = {name: float(df["sum_amount"].sum()) for name, df in base_bids.items()}

    print(f"\nAntall tilbydere: {len(bids)}")
    print(f"Antall poster: {len(set(pd.concat([df['postnr'] for df in base_bids.values()], ignore_index=True)))}")

    print("\nTILBUD (eksklusive opsjoner):")
    print("-" * 80)
    for name, total in sorted(totals.items(), key=lambda x: x[1]):
        option = option_totals.get(name, 0)
        option_str = f" (+ kr {option:,.2f} i opsjoner)" if option > 0 else ""
        print(f"  {name:40s}  kr {total:15,.2f}{option_str}")

    if totals:
        winner = min(totals, key=totals.get)
        winner_total = totals[winner]
        print(f"\n{'🏆 VINNER: ' + winner:40s}  kr {winner_total:15,.2f}")

    # Z-score summary hvis 3+ tilbud
    if len(bids) >= 3:
        print("\nZ-SCORE TOTALER (lavere = bedre):")
        print("-" * 80)

        # Beregn z-score for hver tilbyder
        z_totals = {}
        for name in totals.keys():
            z_totals[name] = 0.0

        # For hver post, beregn z-score
        all_posts = set()
        for df in base_bids.values():
            all_posts.update(df['postnr'].unique())

        for postnr in all_posts:
            post_sums = []
            post_names = []
            for name, df in base_bids.items():
                row = df[df['postnr'] == postnr]
                if not row.empty:
                    post_sums.append(float(row['sum_amount'].iloc[0]))
                    post_names.append(name)

            if len(post_sums) >= 3:
                mean = np.mean(post_sums)
                std = np.std(post_sums, ddof=1)
                if std > 0:
                    for i, name in enumerate(post_names):
                        z = (post_sums[i] - mean) / std
                        z_totals[name] += z

        for name, z_total in sorted(z_totals.items(), key=lambda x: x[1]):
            indicator = "✅" if z_total < -1 else "⚠️" if z_total > 1 else "  "
            print(f"  {indicator} {name:38s}  {z_total:8.2f}")


def print_chapter_summary(base_bids: dict[str, pd.DataFrame], chapter_titles: dict[str, str]):
    """Skriv ut kapitteloppsummering"""
    print("\n" + "=" * 80)
    print("KAPITTELOPPSUMMERING")
    print("=" * 80)

    chapter_df = pd.DataFrame(columns=["kapittel"])
    for name, df in base_bids.items():
        if df.empty:
            continue
        part = df.groupby("kapittel", as_index=False)["sum_amount"].sum().rename(columns={"sum_amount": name})
        if chapter_df.empty:
            chapter_df = part
        else:
            chapter_df = chapter_df.merge(part, on="kapittel", how="outer")

    chapter_df = chapter_df.fillna(0.0)
    chapter_df = chapter_df.sort_values("kapittel")

    for _, row in chapter_df.iterrows():
        kapittel = row["kapittel"]
        kapittel_navn = chapter_titles.get(kapittel, "")

        print(f"\nKapittel {kapittel}: {kapittel_navn}")
        print("-" * 80)

        provider_cols = [col for col in chapter_df.columns if col not in ["kapittel"]]
        sums = {col: float(row[col]) for col in provider_cols}

        if sums and max(sums.values()) > 0:
            winner = min(sums, key=sums.get)
            for name, amount in sorted(sums.items(), key=lambda x: x[1]):
                marker = "🏆 " if name == winner else "   "
                print(f"  {marker}{name:38s}  kr {amount:15,.2f}")


def save_excel(output_path: Path, bids: dict[str, pd.DataFrame], matrix: pd.DataFrame, chapter_df: pd.DataFrame):
    """Lagre resultater til Excel"""
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        # Normalized bids
        for name, df in bids.items():
            df_copy = df.copy()
            if "is_option" in df_copy.columns:
                df_copy = df_copy.drop(columns=["is_option"])
            df_copy.to_excel(writer, index=False, sheet_name=name[:28])

        # Comparison matrix
        if not matrix.empty:
            matrix_copy = matrix.copy()
            if "is_option" in matrix_copy.columns:
                matrix_copy = matrix_copy.drop(columns=["is_option"])
            matrix_copy.to_excel(writer, index=False, sheet_name="Sammenligning")

        # Chapter summary
        if not chapter_df.empty:
            chapter_df.to_excel(writer, index=False, sheet_name="Kapittel")

    print(f"\n✅ Excel-rapport lagret: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Sammenlign anbudstilbud fra kommandolinjen",
        epilog="Eksempel: python bid_compare_cli.py -o rapport.xlsx tilbud1.xlsx tilbud2.xml tilbud3.csv"
    )
    parser.add_argument("files", nargs="+", type=Path, help="Tilbudsfiler (CSV, Excel, eller NS3459 XML)")
    parser.add_argument("-o", "--output", type=Path, default="sammenligning.xlsx", help="Excel-rapport (default: sammenligning.xlsx)")
    parser.add_argument("-v", "--verbose", action="store_true", help="Vis detaljert output med kapitler")

    args = parser.parse_args()

    # Last inn alle filer
    bids = {}
    errors = []

    print("Laster tilbud...")
    for filepath in args.files:
        if not filepath.exists():
            errors.append(f"Filen finnes ikke: {filepath}")
            continue

        try:
            name, df = load_bid_file(filepath)

            # Håndter duplikatnavn
            unique_name = name
            counter = 1
            while unique_name in bids:
                counter += 1
                unique_name = f"{name} ({counter})"

            bids[unique_name] = df
            print(f"  ✓ {filepath.name} -> {unique_name}")
        except Exception as e:
            errors.append(f"Kunne ikke lese {filepath.name}: {e}")
            print(f"  ✗ {filepath.name}: {e}")

    if not bids:
        print("\n❌ Ingen gyldige tilbud ble lastet!")
        if errors:
            print("\nFeil:")
            for error in errors:
                print(f"  - {error}")
        sys.exit(1)

    # Separer base og opsjon
    base_bids = {}
    option_totals = {}
    for name, df in bids.items():
        base_bids[name] = df[df["is_option"] != True].copy()
        option_totals[name] = float(df[df["is_option"] == True]["sum_amount"].sum())

    # Bygg comparison matrix
    matrix = pd.DataFrame(columns=["postnr"])
    sum_columns = []

    for idx, (name, df) in enumerate(base_bids.items()):
        unit_col = f"{name} (enhetspris)"
        sum_col = f"{name} (sum)"
        part = _aggregate_bid_rows(df, unit_col, sum_col)
        if idx == 0:
            matrix = part
        else:
            matrix = matrix.merge(part, on="postnr", how="outer")
        sum_columns.append(sum_col)

    # Kapitteloppsummering
    chapter_titles = _collect_chapter_titles(bids)
    chapter_df = pd.DataFrame(columns=["kapittel"])
    for name, df in base_bids.items():
        if df.empty:
            continue
        part = df.groupby("kapittel", as_index=False)["sum_amount"].sum().rename(columns={"sum_amount": name})
        if chapter_df.empty:
            chapter_df = part
        else:
            chapter_df = chapter_df.merge(part, on="kapittel", how="outer")

    # Vis resultater
    print_summary(bids, base_bids, option_totals)

    if args.verbose and not chapter_df.empty:
        print_chapter_summary(base_bids, chapter_titles)

    # Lagre Excel (alltid)
    save_excel(args.output, bids, matrix, chapter_df)
    print()


if __name__ == "__main__":
    main()
