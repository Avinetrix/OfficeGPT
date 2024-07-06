# ChatGPT Integration for Microsoft Word

This VBA macro allows you to interact with the ChatGPT 4o and 3.5 Turbo models directly from within Microsoft Word. Simply select the text you want to send to ChatGPT, and the macro will retrieve the response and insert it into your document.

## Features

- Send selected text to either the ChatGPT 3.5 or 4o Turbo model
- Receive the response from ChatGPT and insert it into your document
- Robust error handling and debugging capabilities

## Prerequisites

- Microsoft Word with VBA support
- An OpenAI API key (you can obtain one by signing up for an OpenAI account)

## Installation and basic usage

1. Open Microsoft Word and press Alt + F11 to open the Visual Basic Editor.
2. Import the your choice of .bas file your preferred model (File > Import File).
3. Replace `"INSERTAPIKEYHERE"` with your actual OpenAI API key.

The macro(s) can now be used:

4. Select the text you want to send to ChatGPT.
5. Run the `GPT4o` or `GPT35Turbo` macro.
6. The macro will send the selected text to ChatGPT and insert the response on a new line below the selection.

## Recommended usage
For a streamlined experience, add macro button to the Quick Access Toolbar.

1. Right click on the Ribbon and select `Customise Ribbon...`.
2. Filter available options to macros by selecting `Macros` in the `Choose commands from:` dropdown menu.
3. Select your choice of macro.
4. In the right column, click `Home`, click `New Group`, click `Rename...` and rename the group (e.g. ChatGPT).
5. Click `Add >>`, and repeat as required.
6. Select your macro in the right window, click `Rename...`, and rename appropriately (e.g. GPT-3.5 Turbo). Repeat as required.
7. Click `OK` to close the window. 

A new group will appear in the Ribbon's Home tab.

7. Select the text you want to send to ChatGPT.
8. Click the newly added shortcut in the Ribbon. By default, the new group should have added to the far right of the Ribbon, after Add-ins.
9. The macro will send the selected text to ChatGPT and insert the response on a new line below the selection.


## Troubleshooting

If you encounter any issues, check the following:

1. Ensure your API key is correct, has the necessary permissions, and a non-zero credit balance.
2. Check your internet connection.
3. Verify that you have the necessary permissions to make outgoing HTTP requests.

If the problem persists, you can view the full API response in Visual Basic Editor by pressing `Ctrl + G` and running the macro again. This information can assist in diagnosing any remaining issues.

## Contributing

If you find any bugs or have suggestions for improvements, feel free to open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).

## Legal

Use of this repository and its contents (Application) is solely at your own risk. Do not use this Application on mission critical hardware or in sensitive cybersecurity environments. The developer is not responsible for any loss, damage, or disruption caused by the use of this Application, including but not limited to data loss, system failures, hardware damage, or any other consequences, including those arising due to negligence. This Application is provided "as is" without warranty of any kind, express or implied, including but not limited to warranties of merchantability, fitness for a particular purpose, and non-infringement, except as required by Australian law. By using this Application, you agree to these terms and to hold the developer harmless from and indemnify the developer against any and all claims, demands, losses, damages, costs, and expenses (including solicitor's fees) arising out of or in connection with your use of the Application. The developer is additionally not responsible for the creation of black holes, accidental or otherwise.

## Acknowledgement of intellectual property rights

This repository references Microsoft Office and ChatGPT, both of which are the exclusive property of Microsoft Corporation and OpenAI respectively. Microsoft Office, Microsoft Word, and any associated trademarks or logos are the intellectual property of Microsoft Corporation and are used herein solely for informational and referential purposes. ChatGPT, GPT-4o, GPT-3.5 Turbo, and any associated trademarks or logos are the intellectual property of OpenAI and are used herein solely for informational and referential purposes. The developer of this Application is not affiliated with, endorsed by, or in any way officially connected to Microsoft Corporation or OpenAI.
