# Excel QA – Automated Question Answering over Multilingual Excel Spreadsheets

An automated solution that answers natural-language questions about `.xlsx` workbooks (complex layouts, multilingual: EN/VI/ZH/JA/KO/FR). For each question, the system returns **one short, correctly formatted answer** suitable for Exact Match scoring.

- **Input:** a folder of `.xlsx` files and a `queries.json` file.
- **Output:** a `submission.csv` file with two columns, `id` and `answer` (exactly one row per `id`).
- **LLM:** `qwen/qwen3.7-flash` via OpenRouter.

---

## 1. Core idea

Instead of letting the LLM "read and guess" the answer, the system asks the LLM to **write a short Python snippet** that looks up or computes the result from the table data, then **executes that code in a sandbox**. Sums, counts, rankings, rounding and similar operations are performed by the computer rather than estimated by the language model.

### Processing pipeline

```
xlsx ──► Read & normalize (merged cells, header paths, multi-table, formulas)
     ──► Text dump (with cell addresses)
     ──► (Optional) Routing: select relevant sheets for large workbooks
     ──► LLM generates Python code ──► Sandbox runs it ──► (on error → one repair attempt)
     ──► Multi-vote (VOTES) ──► Judge LLM when votes disagree
     ──► Post-processing per the format requested in the question (rounding, dates, Yes/No, labels…)
     ──► submission.csv
```

### Key components

| Component | Description |
|---|---|
| **Excel → `Sheet` reader** | Handles merged cells, multi-level headers (header paths like `A > B`), automatically splits multiple stacked tables within one sheet, and cleans values (missing markers such as `-` and `n/a`, accounting-style numbers, full-width characters, etc.). |
| **Formula evaluator** | For formula cells without a cached value, the system computes them itself (supports `SUM`, `IF`, `VLOOKUP`, `INDEX/MATCH`, `SUMIF(S)`, `COUNTIF(S)`, `ROUND*`, … including cross-sheet references and external-link caches). |
| **Cell lookup by address** | Questions such as "what does cell D5 contain / what is its formula" are answered directly from the source file (`CELL LOOKUP`), without calling the LLM. |
| **Sandbox** | Runs LLM-generated code with a restricted set of builtins and modules (`math`, `re`, `datetime`, `statistics`, …) and a timeout. |
| **Voting + Judge** | Generates several candidate answers (3 by default); if they disagree, an LLM judge compares the code and the table to pick the best answer. |
| **Multilingual post-processing** | Detects rounding/format requirements in the question (`2 decimal places`, `保留2位小数`, `소수점 2자리`, `round up`, `làm tròn xuống`, …), formats dates and Yes/No answers, and restores original labels exactly (preserving dots and commas). |
| **Sheet routing** | For large workbooks (many sheets, long dumps), the LLM selects only the necessary sheets to reduce tokens and latency. |

---

## 2. Repository structure

```
.
├── excelqa.py          # main source code
├── requirements.txt    # dependencies
├── Dockerfile
├── LICENSE
└── README.md
```

Dependencies are listed in `requirements.txt` (`openpyxl`, `requests`, `pandas`, `numpy`, `xlrd`, `python-dateutil`, `scipy`).

> The system does **not** use local embedding or reranker models, so memory usage stays low.

---

## 3. Data format

**`queries.json`** – a list of questions, each containing:

```json
[
  {
    "id": "q001",
    "file_name": "report.xlsx",
    "query": "What is the total revenue in 2023? Round to 2 decimal places."
  }
]
```

**`submission.csv`** – the result:

```csv
id,answer
q001,12345.67
```

---

## 4. Usage

### 4.1. Run with Docker (recommended)

Build the image:

```bash
docker build -t excelqa .
```

Run:

```bash
docker run --rm \
  -e OPENROUTER_API_KEY="<your-api-key>" \
  -v /path/to/data:/data:ro \
  -v /path/to/out:/out \
  excelqa \
  --xlsx-dir /data \
  --queries /data/queries.json \
  --output /out/submission.csv
```

The image `ENTRYPOINT` is `python excelqa.py`, so you only need to pass the arguments after the image name.

### 4.2. Run directly with Python

```bash
pip install -r requirements.txt
export OPENROUTER_API_KEY="<your-api-key>"

python excelqa.py \
  --xlsx-dir ./data \
  --queries ./data/queries.json \
  --output ./submission.csv
```

### 4.3. Command-line arguments

| Argument | Description |
|---|---|
| `--xlsx-dir` | Folder containing the `.xlsx` files (searched recursively). |
| `--queries` | Path to `queries.json`. If omitted, `queries.json` is looked up in the data folder. |
| `--output`, `-o` | Path of `submission.csv` (default: `submission.csv`). |
| `--openrouter-key` | API key (or use the `OPENROUTER_API_KEY` environment variable). |
| `--openrouter-model` | Model name (default: `qwen/qwen3.7-flash`). |
| `--openrouter-base-url` | Base URL, e.g. `https://openrouter.ai/api/v1`. |
| `--input`, `-i` | (Legacy) folder containing both `queries.json` and the `.xlsx` files. |
| `--api-key` | (Legacy) equivalent to `--openrouter-key`. |
| `--dump <file.xlsx>` | Print the text dump of one xlsx file and exit (useful for checking the data-reading stage). |

> No paths are hardcoded; the program runs unchanged on hidden test sets.

---

## 5. Configuration via environment variables

All parameters have sensible defaults; set them only when tuning is needed.

**Connection & resources**

| Variable | Default | Description |
|---|---|---|
| `OPENROUTER_API_KEY` | – | **Required.** OpenRouter API key (never written into the source code). |
| `OPENROUTER_MODEL` | `qwen/qwen3.7-flash` | Model to use. |
| `OPENROUTER_API_URL` | `https://openrouter.ai/api/v1/chat/completions` | API endpoint. |
| `WORKERS` | `6` | Number of questions processed in parallel. |
| `TIME_BUDGET_SEC` | `2100` | Overall time budget in seconds, kept under the 40-minute limit. |
| `MAX_COST_PER_Q` | `0.08` | Per-question cost cap (USD), below the allowed 0.10 USD. |
| `LLM_MAX_TOKENS` | `3000` | Maximum output tokens per call. |
| `PRICE_IN_PER_M` / `PRICE_OUT_PER_M` | `0.03` / `0.13` | Unit prices (USD per 1M tokens) used for cost estimation. |
| `HTTP_TIMEOUT`, `LLM_RETRIES` | `90`, `4` | Timeout and number of retries for API calls. |
| `REASONING` | `off` | `off` disables the model's reasoning mode to save tokens and time. |

**Voting & judge**

| Variable | Default | Description |
|---|---|---|
| `VOTES` | `3` | Number of votes per question. |
| `VOTE_TEMP` | `0.3` | Temperature for votes after the first (the first vote uses 0). |
| `PARALLEL_VOTES` | `1` | Run the first two votes in parallel. |
| `JUDGE` | `1` | Enable/disable the judge when votes disagree. |
| `JUDGE_OVERRIDE` | `1` | `0` = call the judge only when there is no majority. |
| `REJECT_ERROR_ANSWERS` | `1` | Do not accept strings like "Error…" as answers. |

**Data reading & routing**

| Variable | Default | Description |
|---|---|---|
| `MULTI_TABLE` | `1` | Automatically split multiple stacked tables in a sheet. |
| `ROUTING` | `1` | Enable relevant-sheet selection. |
| `ROUTE_MIN_CHARS` | `12000` | Only route when the dump is longer than this threshold. |
| `ROUTE_SHRINK_RATIO` | `0.8` | Only use the routing result if it shrinks the dump below this ratio. |
| `MAX_HEADER_ROWS` | `12` | Maximum number of rows treated as headers. |
| `CELL_FALLBACK` | `1` | Look up cells by address and read raw formulas. |
| `SANDBOX_TIMEOUT` | `15` | Timeout (seconds) when running LLM-generated code. |

**Logging**

| Variable | Default | Description |
|---|---|---|
| `DEBUG_LOG` | `debug_log.jsonl` | Detailed per-question log (votes, code, judge, tokens, cost, time). |
| `LOG_CANDS` | `1` | Also log the code and answer of every vote. |

---

## 6. Meeting the task constraints

| Constraint | How it is met |
|---|---|
| Model `qwen/qwen3.7-flash` via OpenRouter | Uses exactly this model by default, called over HTTP with `requests`. |
| No hardcoded API key | The key comes from the `OPENROUTER_API_KEY` environment variable or a command-line argument. |
| RAM ≤ 1 GB | No local embedding/reranker; each workbook is cached once and holds only text data. |
| Package ≤ 100 MB | Only `excelqa.py`, `requirements.txt`, `Dockerfile` and `README.md`. |
| ≤ 0.10 USD per question | A per-question cost counter stops processing when `MAX_COST_PER_Q` (default 0.08) is exceeded. |
| ≤ 40 minutes for 20 questions | Questions run in parallel, with `TIME_BUDGET_SEC` as a hard deadline for the whole pipeline. |
