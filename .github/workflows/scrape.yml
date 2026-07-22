name: Art-Gift to Temu

on:
  workflow_dispatch:
    inputs:
      category:
        description: "Category to scan"
        required: true
        type: choice
        default: "all"
        options:
          - all
          - tshirts
          - bodysuits
      max_products:
        description: "10-15 for test; 0 for all products"
        required: true
        default: "15"
      shipping_template:
        description: "Exact Temu shipping template name"
        required: false
        default: ""

jobs:
  scrape:
    runs-on: ubuntu-latest
    timeout-minutes: 360
    steps:
      - uses: actions/checkout@v4

      - uses: actions/setup-python@v5
        with:
          python-version: "3.11"
          cache: pip

      - name: Install dependencies
        run: |
          python -m pip install --upgrade pip
          pip install -r requirements.txt
          python -m playwright install --with-deps chromium

      - name: Prepare configuration
        env:
          CATEGORY: ${{ inputs.category }}
          MAX_PRODUCTS: ${{ inputs.max_products }}
          SHIPPING_TEMPLATE: ${{ inputs.shipping_template }}
        run: |
          python - <<'PY'
          import json, os
          cfg = json.load(open('config.example.json', encoding='utf-8'))
          selected = os.environ.get('CATEGORY', 'all')
          slug_by_input = {'tshirts': 'тениски', 'bodysuits': 'бодита'}
          if selected in slug_by_input:
              wanted = slug_by_input[selected]
              cfg['categories'] = [c for c in cfg['categories'] if c.get('slug') == wanted]
          cfg['max_products'] = int(os.environ.get('MAX_PRODUCTS') or 15)
          cfg['shipping_template'] = os.environ.get('SHIPPING_TEMPLATE', '')
          json.dump(cfg, open('config.json','w',encoding='utf-8'), ensure_ascii=False, indent=2)
          print('Selected category:', selected)
          print('Configured category slugs:', [c['slug'] for c in cfg['categories']])
          PY

      - name: Run scraper
        run: python artgift_scraper.py --config config.json

      - name: Upload output
        if: always()
        uses: actions/upload-artifact@v4
        with:
          name: artgift-temu-output-${{ inputs.category }}
          path: output/
          if-no-files-found: warn
          retention-days: 14
