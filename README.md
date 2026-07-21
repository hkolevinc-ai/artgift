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
