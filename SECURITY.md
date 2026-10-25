# Security Policy

## Scope

ppt-factory is a local CLI. It sends a topic, or text extracted from a `.pptx` you supply, to Amazon
Bedrock and renders the model's JSON response into a new `.pptx`. It runs no server and stores no
credentials.

Areas with a real security surface:

- **AWS credentials.** boto3 reads your standard AWS credential chain. Nothing in this repo should
  ever contain keys. Use least-privilege IAM: `bedrock:InvokeModel` on the one model, nothing broader.
- **Content you send to Bedrock.** `--from-pptx` sends the full slide text and the first 300
  characters of each slide's speaker notes. Don't convert decks you're not permitted to share
  with your AWS account's model provider.
- **Model output.** The response is parsed as JSON and rendered. Reports about unsafe handling of
  that output are in scope.

## Reporting a vulnerability

Please use GitHub's private vulnerability reporting on this repo
(**Security → Report a vulnerability**). Do not open a public issue for security findings.

Expect an acknowledgement within 7 days. This is a personal project maintained by one person, so
fixes are best effort.

## Supported versions

Only the `main` branch is maintained.
