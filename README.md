Assistant for Word (AssistantForWord)

Overview

AssistantForWord is a VSTO Word add-in that integrates with OpenAI to provide AI-assisted text operations directly inside Microsoft Word. The add-in provides a custom Ribbon, a side panel for settings and prompts, and quick actions that apply prompts to selected text and insert AI responses back into the document.

Key features

- Custom Ribbon with dynamic prompt buttons
- Side panel for configuration (model, temperature, token size) and prompt management
- Uses OpenAI (via OpenAI-API NuGet package) to create chat completions
- Saves configuration to AppData (AIAssitant\\Config.dll)

Requirements

- Visual Studio (tested with Visual Studio Enterprise 2026)
- Microsoft Office Word (VSTO runtime)
- .NET Framework 4.7.2 (project target)
- NuGet packages restored (see packages.config in the project)

Build & run

1. Open AssistantForWord.sln in Visual Studio.
2. Restore NuGet packages.
3. Build the solution (set AssistantForWord as the startup project if needed).
4. Run (F5) to debug the add-in; Word will start with the add-in loaded in the debug session.

Configuration

- The add-in stores settings in %APPDATA%\\AIAssitant\\Config.dll (serialized AssistantConfig object).
- Use the add-in Settings pane (open the Assistant pane from the Ribbon) to set your OpenAI API key, model, token size, temperature, and manage prompt templates.

Important security notes

- Current code contains a hard-coded OpenAI API key in AssistantForWord\\OpenAIClient.cs. This is insecure and must be removed before publishing or sharing the repository. Replace the hard-coded key with configuration-based usage, e.g.:

  var client = new OpenAIAPI(config.APIKEY);

- The project uses BinaryFormatter to serialize configuration to disk (ProcessData.cs). BinaryFormatter is insecure for untrusted data. Consider replacing it with a safer serializer (e.g., JSON via System.Text.Json or Newtonsoft.Json) and a different file extension.

- Do not commit real API keys or other secrets to source control. Use environment variables, secure stores, or configuration encrypted at rest for production deployments.

Contributing

- Fork the repository, create a topic branch, and submit pull requests targeting main.
- For fixes involving secrets or serialization, add migration steps and tests where possible.

License

- No license is included in this repository. Add a LICENSE file or update this README with the chosen license.

Contact

- For questions about building or running the add-in, open an issue in the repository.

