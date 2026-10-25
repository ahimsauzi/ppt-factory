# ppt-factory

AI-powered presentation builder using **claude-opus-4-6 via Amazon Bedrock**.
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
└── .gitignore
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
AWS Console → Amazon Bedrock → Model access → Request access → Anthropic Claude claude-opus-4-6
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

