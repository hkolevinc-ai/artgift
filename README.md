# Art-Gift → Temu scraper

The scraper scans only these two Art-Gift categories:

- `https://art-gift.net/тениски`
- `https://art-gift.net/бодита`

It collects every product, creates a separate row for each offered **colour × size** combination, links all SKU rows through the same **Contribution Goods** value, converts/reads the price in **EUR**, and marks personalization products as:

- `Custom product`
- `Single technique`
- `Leather/fabric customization technique`
- `Digital printing`

The actual product personalization fields (for example desired name and additional changes) are copied into the product bullets/description.

## Before the full run

1. Copy `config.example.json` to `config.json`.
2. Replace `shipping_template` with the exact Temu shipping-template name from the seller account.
3. Add `eu_responsible_person` when required for the account/product setup.
4. Verify the default package weight and dimensions. They are working assumptions, not measurements from Art-Gift.
5. Review the Temu category for unisex bodysuits. The supplied template has separate Baby Boys/Baby Girls categories, so unisex bodysuits default to `26402` and obvious girl-specific products use `26325`.

The example configuration limits the first run to 15 products. Set `max_products` to `0` for the complete categories after reviewing the test output.

## Local run

```bash
python -m pip install -r requirements.txt
python -m playwright install chromium
cp config.example.json config.json
python artgift_scraper.py --config config.json
```

The files are written to `output/`. Because the Temu template supports about 2,000 data rows, the scraper automatically creates `part01`, `part02`, etc.

## GitHub Actions

Upload the whole folder to a GitHub repository and run **Actions → Art-Gift to Temu → Run workflow**.

Workflow inputs:

- `max_products=15` for the first test;
- `max_products=0` for the complete run.

The result is downloaded as the `artgift-temu-output` artifact.

## Important behaviour

- The scraper first uses a browser to click the colour/size selectors and capture exact offered combinations.
- If the storefront JavaScript changes or the browser cannot enumerate them, it falls back to the visible colour × size cross-product and records a warning.
- Product prices are read from the site's EUR price field/structured data; no BGN conversion is performed.
- Existing workbook formatting, formulas, dropdowns and validations are preserved through direct XLSX package editing.
- A `run_report.json` file lists counts, output parts, failures and warnings.

## Separate category tests

The GitHub Actions form now has a **Category to scan** selector:

- `tshirts` — scans only `Тениски`;
- `bodysuits` — scans only `Бодита`;
- `all` — scans both categories.

For validation, run `tshirts` with `max_products=10`, then `bodysuits` with
`max_products=10`. The generated `run_report.json` includes
`products_by_category`, so it is immediately clear which category was scanned.

The browser variation reader also waits for the selected colour and its price to
settle before exporting the SKU rows. This prevents a previous colour's price
from being reused during slow storefront updates.

## Bodysuit variation handling (V3)

Art-Gift bodysuits do not use the same selector layout as T-shirts. On bodysuits,
the storefront's main selector is **sleeve length**, while the secondary selector
contains the actual garment sizes (56, 62, 68, etc.). V3 handles this separately:

- Temu **Size** receives the real garment size;
- Temu **Color** is set to `White` because the tested bodysuits do not offer a colour selector;
- short-sleeve and long-sleeve versions become separate **Contribution Goods**;
- every garment size remains a separate **Contribution SKU**;
- the price is read after both size and sleeve have been selected, so the long-sleeve surcharge is preserved.

This split is necessary because the supplied Temu categories expose Size and Color
as SKU-level sale properties, while sleeve length is a product-level property.
