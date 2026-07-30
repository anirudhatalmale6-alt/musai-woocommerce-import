#!/usr/bin/env python3
"""
Convert MUSAI 'PROD ZDS.xlsx' clothing catalogue into a WooCommerce-native-importer CSV.

Grouping:  one 'variable' parent per MODELCODE, one 'variation' row per PRODUCTCODE (SKU).
Attributes:
  - Size   (global pa_size,   used for variations)
  - Color  (global pa_color,  used for variations)
  - Gender (global pa_gender, product-level, NOT used for variations)
  - Dimensions -> stored per-variation as a custom field  (Meta: dimensions)
                  and as a visible parent custom attribute listing all measures.
Images: featured + gallery on parent (unique model/product shots per colour),
        one image per variation. Importer sideloads the URLs.

Single-variation models are emitted as 'simple' products.
"""
import csv, sys, openpyxl
from collections import defaultdict, OrderedDict

SRC = "PROD ZDS.xlsx"

# WooCommerce native importer header
HEADER = [
    "Type","SKU","Name","Published","Is featured?","Visibility in catalog",
    "Short description","Description","Tax status","In stock?","Backorders allowed?",
    "Sold individually?","Regular price","Categories","Tags","Images","Parent",
    "Position",
    "Attribute 1 name","Attribute 1 value(s)","Attribute 1 visible","Attribute 1 global","Attribute 1 default",
    "Attribute 2 name","Attribute 2 value(s)","Attribute 2 visible","Attribute 2 global","Attribute 2 default",
    "Attribute 3 name","Attribute 3 value(s)","Attribute 3 visible","Attribute 3 global","Attribute 3 default",
    "Attribute 4 name","Attribute 4 value(s)","Attribute 4 visible","Attribute 4 global","Attribute 4 default",
    "Meta: measures","Meta: ean",
]

# column indexes in the source sheet
C_SKU,C_MEAS,C_MODEL,C_NAME,C_DESC,C_COMP,C_OBS,C_MAKER,C_GENDER,C_SIZE,\
C_COLOR,C_COLORCODE,C_PRODIMG,C_MODIMG,C_WASH,C_PRICE,C_EAN,C_CAT = range(18)

# fallback names for the one model that ships without a MODELNAME
NAME_FALLBACK = {"MU0400-K": "Jachetă - Maji Copii"}


def clean(v):
    return "" if v is None else str(v).strip()


def dedup(seq):
    return list(OrderedDict.fromkeys(x for x in seq if x))


# WooCommerce importer splits multi-value attribute cells on COMMA. Size/Color/Gender
# contain no commas, so we join with ", ". Measures use EU decimals ("11,5cm") -> the
# comma would wrongly split them, so normalise to "." for the display attribute only
# (the per-variation Meta: measures keeps the original value untouched).
def dim_display(m):
    return m.replace(",", ".")


def measure_pretty(m):
    """'[47cm]·[68cm]' -> '47 cm · 68 cm' for human-readable display."""
    return (m.replace("[", "").replace("]", "")
             .replace("cm", " cm").replace("·", " · ").replace("  ", " ").strip())


def measures_table(size_measure):
    """HTML size -> measurement table for the product description."""
    if not size_measure:
        return ""
    rows = "".join(
        f"<tr><td style='padding:4px 12px;border:1px solid #ddd'>{s}</td>"
        f"<td style='padding:4px 12px;border:1px solid #ddd'>{measure_pretty(m)}</td></tr>"
        for s, m in size_measure.items()
    )
    return (
        "<h4>Măsurători / Measurements</h4>"
        "<table style='border-collapse:collapse;margin:8px 0'>"
        "<tr><th style='padding:4px 12px;border:1px solid #ddd;text-align:left'>Mărime / Size</th>"
        "<th style='padding:4px 12px;border:1px solid #ddd;text-align:left'>Măsurători</th></tr>"
        f"{rows}</table>"
    )


def build_description(r):
    """Assemble an HTML product description from description / composition / observation."""
    parts = []
    desc = clean(r[C_DESC]).replace("·", "<br>·")
    if desc:
        parts.append(f"<p>{desc}</p>")
    comp = clean(r[C_COMP])
    if comp:
        parts.append(f"<p><strong>Compoziție:</strong> {comp.lstrip('· ').strip()}</p>")
    obs = clean(r[C_OBS])
    if obs:
        parts.append(f"<p>{obs.replace('·', '<br>·')}</p>")
    return "".join(parts)


def load_models():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = [r for r in ws.iter_rows(values_only=True)][1:]
    rows = [r for r in rows if r and clean(r[C_SKU])]
    models = OrderedDict()
    for r in rows:
        models.setdefault(clean(r[C_MODEL]), []).append(r)
    return models


def emit_model(w, model_code, variations):
    first = variations[0]
    name = clean(first[C_NAME]) or NAME_FALLBACK.get(model_code, model_code)
    categories = clean(first[C_CAT])
    description = build_description(first)

    sizes = dedup(clean(v[C_SIZE]) for v in variations)
    colors = dedup(clean(v[C_COLOR]) for v in variations)
    measures = dedup(clean(v[C_MEAS]) for v in variations)
    genders = dedup(clean(v[C_GENDER]) for v in variations)
    gender = ", ".join(genders)

    # measures depend on SIZE only -> one measurement per size (verified in source data)
    size_measure = OrderedDict()
    for v in variations:
        s, m = clean(v[C_SIZE]), clean(v[C_MEAS])
        if s and m and s not in size_measure:
            size_measure[s] = m
    # surface the size guide in the product description
    description = description + measures_table(size_measure)

    # gallery: model shots + product shots, unique, in order of appearance
    gallery = dedup([clean(v[C_MODIMG]) for v in variations] +
                    [clean(v[C_PRODIMG]) for v in variations])

    simple = len(variations) == 1

    if simple:
        v = variations[0]
        w.writerow({
            "Type": "simple", "SKU": clean(v[C_SKU]), "Name": name,
            "Published": 1, "Is featured?": 0, "Visibility in catalog": "visible",
            "Short description": "", "Description": description,
            "Tax status": "taxable", "In stock?": 1, "Backorders allowed?": 0,
            "Sold individually?": 0, "Regular price": clean(v[C_PRICE]),
            "Categories": categories, "Tags": clean(v[C_MAKER]),
            "Images": ",".join(gallery), "Parent": "", "Position": 0,
            "Attribute 1 name": "Size",  "Attribute 1 value(s)": clean(v[C_SIZE]),  "Attribute 1 visible": 1, "Attribute 1 global": 1, "Attribute 1 default": "",
            "Attribute 2 name": "Color", "Attribute 2 value(s)": clean(v[C_COLOR]), "Attribute 2 visible": 1, "Attribute 2 global": 1, "Attribute 2 default": "",
            "Attribute 3 name": "Gender","Attribute 3 value(s)": gender,            "Attribute 3 visible": 1, "Attribute 3 global": 1, "Attribute 3 default": "",
            "Attribute 4 name": "Measures","Attribute 4 value(s)": dim_display(clean(v[C_MEAS])),"Attribute 4 visible": 1, "Attribute 4 global": 0, "Attribute 4 default": "",
            "Meta: measures": clean(v[C_MEAS]), "Meta: ean": clean(v[C_EAN]),
        })
        return

    # variable parent
    w.writerow({
        "Type": "variable", "SKU": model_code, "Name": name,
        "Published": 1, "Is featured?": 0, "Visibility in catalog": "visible",
        "Short description": "", "Description": description,
        "Tax status": "taxable", "In stock?": 1, "Backorders allowed?": 0,
        "Sold individually?": 0, "Regular price": "",
        "Categories": categories, "Tags": clean(first[C_MAKER]),
        "Images": ",".join(gallery), "Parent": "", "Position": 0,
        "Attribute 1 name": "Size",  "Attribute 1 value(s)": ", ".join(sizes),  "Attribute 1 visible": 1, "Attribute 1 global": 1, "Attribute 1 default": "",
        "Attribute 2 name": "Color", "Attribute 2 value(s)": ", ".join(colors), "Attribute 2 visible": 1, "Attribute 2 global": 1, "Attribute 2 default": "",
        "Attribute 3 name": "Gender","Attribute 3 value(s)": gender,            "Attribute 3 visible": 1, "Attribute 3 global": 1, "Attribute 3 default": "",
        "Attribute 4 name": "Measures","Attribute 4 value(s)": ", ".join(dim_display(m) for m in measures),"Attribute 4 visible": 1, "Attribute 4 global": 0, "Attribute 4 default": "",
        "Meta: measures": "", "Meta: ean": "",
    })

    # variations
    for i, v in enumerate(variations, start=1):
        meas = clean(v[C_MEAS])
        # variation description shows under the dropdowns when that size is selected
        var_desc = f"Măsurători: {measure_pretty(meas)}" if meas else ""
        w.writerow({
            "Type": "variation", "SKU": clean(v[C_SKU]), "Name": "",
            "Published": 1, "Is featured?": 0, "Visibility in catalog": "visible",
            "Short description": "", "Description": var_desc,
            "Tax status": "taxable", "In stock?": 1, "Backorders allowed?": 0,
            "Sold individually?": 0, "Regular price": clean(v[C_PRICE]),
            "Categories": "", "Tags": "",
            "Images": clean(v[C_PRODIMG]), "Parent": model_code, "Position": i,
            "Attribute 1 name": "Size",  "Attribute 1 value(s)": clean(v[C_SIZE]),  "Attribute 1 visible": "", "Attribute 1 global": 1, "Attribute 1 default": "",
            "Attribute 2 name": "Color", "Attribute 2 value(s)": clean(v[C_COLOR]), "Attribute 2 visible": "", "Attribute 2 global": 1, "Attribute 2 default": "",
            "Attribute 3 name": "", "Attribute 3 value(s)": "", "Attribute 3 visible": "", "Attribute 3 global": "", "Attribute 3 default": "",
            "Attribute 4 name": "", "Attribute 4 value(s)": "", "Attribute 4 visible": "", "Attribute 4 global": "", "Attribute 4 default": "",
            "Meta: measures": meas, "Meta: ean": clean(v[C_EAN]),
        })


def write_csv(path, models):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=HEADER, extrasaction="ignore")
        w.writeheader()
        for mc, variations in models.items():
            emit_model(w, mc, variations)


def main():
    models = load_models()
    if len(sys.argv) > 1 and sys.argv[1] == "test":
        pick = ["MU0204", "MU0602", "MU0606"]
        subset = OrderedDict((k, models[k]) for k in pick if k in models)
        write_csv("MUSAI_woocommerce_TEST.csv", subset)
        n = sum(len(v) for v in subset.values())
        print(f"TEST written: {len(subset)} products, {n} variation rows -> MUSAI_woocommerce_TEST.csv")
    else:
        write_csv("MUSAI_woocommerce_FULL.csv", models)
        n = sum(len(v) for v in models.values())
        print(f"FULL written: {len(models)} products, {n} variation rows -> MUSAI_woocommerce_FULL.csv")


if __name__ == "__main__":
    main()
