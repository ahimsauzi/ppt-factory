# ppt-factory

A Python CLI that turns a topic, or an existing deck, into a branded executive presentation using
Claude on Amazon Bedrock. Every plan the model returns is validated before anything is rendered.
Single script — no package install required.

---

## Project Structure

```
ppt-factory/
├── factory.py               # single script — run this directly
├── requirements.txt         # boto3, python-pptx
├── requirements-dev.txt     # + pytest, ruff
├── tests/
│   └── test_ppt_factory.py  # pytest suite (no AWS credentials required)
├── scripts/
│   └── new_theme.py         # helper: scaffold a new theme
├── .github/workflows/ci.yml # lint + tests on every push and PR
├── SECURITY.md
└── LICENSE
```

---

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. AWS credentials

If you have already run `aws configure`, you are ready — no extra steps needed.
boto3 uses `~/.aws/credentials` automatically.

```bash
# Verify
aws sts get-caller-identity
aws configure get region
```

### 3. Enable Claude in Bedrock

```
AWS Console → Amazon Bedrock → Model access → Request access → Anthropic Claude
```

Then point ppt-factory at the model or inference profile you enabled:

```bash
export PPT_FACTORY_MODEL_ID="<your Bedrock model or inference profile ID>"
```

### 4. IAM permissions required (minimum)

```json
{
  "Version": "2012-10-17",
  "Statement": [{
    "Effect": "Allow",
    "Action": "bedrock:InvokeModel",
    "Resource": "arn:aws:bedrock:us-east-1::foundation-model/anthropic.claude-opus-4-6*"
  }]
}
```

---

## Usage

Run all commands from the repo root where `factory.py` lives.

### Generate a new deck from a topic

```bash
python factory.py --topic "Zero Trust for Financial Services" --theme midnight
python factory.py --topic "Bedrock Model Deployment" --audience executive --slides 10 --theme teal
```

### Convert an existing PPTX to a new theme

Extracts all content, sends to Bedrock for reinterpretation, rebuilds with the target theme.

```bash
python factory.py --from-pptx existing_deck.pptx --theme meridian
python factory.py --from-pptx existing_deck.pptx --theme midnight --output rethemed.pptx
```

### Review the plan before building (saves an API call on re-runs)

```bash
# Generate plan only — no .pptx built, no second API call
python factory.py --topic "AWS Security Hub" --plan-only > plan.json

# Build from saved plan — no API call
python factory.py --from-json plan.json --theme forest
python factory.py --from-json plan.json --theme teal
```

---

## Validation

The model returns a JSON plan, and nothing is rendered until that plan passes `validate_plan()`:

- **Structure:** a title, 1-30 slides, each with a known layout, starting with `title` and ending with `closing`.
- **Fields:** required text present and non-empty; optional text, if present, is a string; `phase_color` is one of `accent_1`-`accent_3`.
- **Capacity:** list lengths stay within what each layout can draw (for example, at most 3 timeline phases and 5 items per panel). Content that would be cut off is an error, not a silent drop.

If a Bedrock response fails validation or isn't valid JSON, the model gets **one repair attempt**, with the
specific errors fed back. If that also fails, the CLI exits with code 2 and writes nothing. Plans loaded
with `--from-json` and plans passed to `build_pptx()` are validated the same way.

---

## Themes

| Name | Style |
|------|-------|
| `meridian` | Deep navy / crimson — house style (default) |
| `midnight` | Dark slate / electric blue — enterprise tech |
| `teal` | Deep teal / seafoam — trust & security |
| `corporate` | Charcoal / professional grey |
| `forest` | Deep green / sage — sustainability |

---

## Running Tests

No AWS credentials required.

```bash
pip install -r requirements-dev.txt
pytest tests/
```

`--plan-only` prints only the JSON plan to stdout. Progress messages go to stderr, so
`> plan.json` produces a clean file.

---

## Security

ppt-factory sends your topic, or the text of a deck you convert, to Amazon Bedrock under your own AWS
account. Use a least-privilege IAM policy like the one above, and keep credentials in your AWS
credential chain, never in this repo. See [SECURITY.md](SECURITY.md) for scope and reporting.

## License

[MIT](LICENSE)
