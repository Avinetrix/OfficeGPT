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
Add a macro button to the Quick Access Toolbar

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
