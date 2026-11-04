import glob
import os

from extracttab import find_tables, open_pdf, write_workbook


def main():
    all_tables = []
    for path in sorted(glob.glob("samples/*.pdf")):
        pdf = open_pdf(path)
        tables = find_tables(pdf)
        print(f"\n{os.path.basename(path)}")
        print(f"  frag fonts: {sorted(pdf.frag_fonts)[:4]} ...")
        print(f"  tables: {len(tables)}")
        for t in tables:
            print(f"    {t.label:16} p{t.page:<3} {t.nrows}x{t.ncols:<3} "
                  f"suspect={t.suspect_count} num={t.numeric_ratio:.0%}")
        all_tables.extend(tables)
        pdf.close()

    print(f"\n===== total {len(all_tables)} tables =====")
    target = [t for t in all_tables if t.source_name.endswith("0061514.pdf")]
    if target:
        t = target[0]
        print(f"\n验证 {t.label} ({t.source_name}) 前 3 行:")
        for row in t.data[:3]:
            print("   " + " | ".join(f"{c.text[:18]:18}" for c in row))

    out = write_workbook(all_tables, "out_test.xlsx", with_source=True)
    print(f"\nwrite -> {out}  ({out.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    main()
