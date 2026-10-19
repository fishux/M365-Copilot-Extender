# M365 Copilot Extender

A new Chrome Manifest V3 extension project for manual model selection in [Microsoft 365 Copilot](https://m365.cloud.microsoft).

Status: research and repository setup only. Extension implementation will follow the project owner's next instructions.

The intended target is Google Chrome on Windows, loaded as an unpacked extension. Model selection uses reverse-engineered Copilot behavior rather than an official Microsoft API. A requested model must not be presented as a verified backend model merely because its outgoing `tone` was changed.

## Technical references

- [cramt/m365-copilot-proxy](https://github.com/cramt/m365-copilot-proxy): protocol, model tones, routing, response metadata, and quota research.
- [aweussom/i-want-my-model-picker-back](https://github.com/aweussom/i-want-my-model-picker-back): picker research and WebSocket interception experiments.
- [Research baseline](docs/research.md): source revisions, findings, and outstanding live checks.

No upstream implementation has been incorporated.
