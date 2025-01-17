# VBA Monologger website documentation

> VBA Monologger is an advanced and flexible logging solution for VBA (*Visual Basic for Applications*) ecosystem. It is largely inspired by the [Monolog](https://github.com/Seldaek/monolog) library in PHP, which itself is inspired by the [Logbook](https://logbook.readthedocs.io/en/stable/) library in Python.
>
> Website documentation : https://6i-software.github.io/vba-monologger/


## Local installation

1. Create a Python virtual environment into the folder `./venv`, depending on the Python version you are running.
    ```
    > python -m venv venv
    ```
   
2. Activate your Python virtual environment.
   ```
   > .\venv\Scripts\activate
    ```
   
3. Install dependencies into venv
    ```
    (venv)> pip install -r requirements.txt
    ```

4. Start a local server. It runs a local web server on your machine, usually accessible at http://127.0.0.1:20100.
   ```
   mkdocs serve
   ```

5. Build the website. This build command is used to generate your full documentation site as static HTML files.
   ```
   mkdocs build
   ```


## Github pipeline CI/CD

```yaml
# ./.github/workflows/ci.yml

name: CI

# Controls when the workflow will run
on:
  push:
    branches: ["documentation"]
  pull_request:
    branches: ["documentation"]
    
  # Allows you to run this workflow manually from the Actions tab
  workflow_dispatch:

permissions:
  contents: write    
  
jobs:
  deploy:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - name: Configure Git Credentials
        run: |
          git config user.name github-actions[bot]
          git config user.email github-actions[bot]@users.noreply.github.com
      - uses: actions/setup-python@v5
        with:
          python-version: 3.x
      - run: echo "cache_id=$(date --utc '+%V')" >> $GITHUB_ENV 
      - uses: actions/cache@v4
        with:
          key: mkdocs-material-${{ env.cache_id }}
          path: .cache
          restore-keys: |
            mkdocs-material-
      - run: pip install mkdocs-material 
      - run: mkdocs gh-deploy --force --config-file ./mkdocs.yml
```