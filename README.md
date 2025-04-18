# Writing Planner Advanced Add-in for MS Word

## Overview

The Writing Planner Advanced is an add-in for Microsoft Word that enhances the document creation process through a two-staged AI inference approach. It integrates seamlessly into MS Word, providing users with a powerful tool to plan and generate documents efficiently.

## Features

1. **Two-Staged AI Inference**: 
   - Stage 1: Creates a document structure along with required prompts.
   - Stage 2: Allows users to generate base text for each section separately using AI.

2. **BYOK (Bring Your Own API Key)**: 
   - The add-in relies on the OpenRouter API for AI inference.
   - Users are required to provide their own API key for usage.

## Technical Details

- **Integration**: The add-in docks inside MS Word, providing a user-friendly interface within the familiar Word environment.
- **AI Inference**: Utilizes the OpenRouter API for generating document structures and content.
- **Customization**: Allows users to customize the document planning process according to their needs.

## Development

The project is developed using React and utilizes various components from `@fluentui/react`. The main application logic is contained within the `App.js` file located in `src/taskpane/components/`.

## Setup and Usage

1. **API Key**: Obtain an API key from OpenRouter.
2. **Installation**: Follow the installation instructions provided in the project documentation or repository.
3. **Usage**: 
   - Open MS Word and access the Writing Planner Advanced add-in.
   - Follow the in-app instructions to plan and generate your document.

## Contributing

Contributions to the Writing Planner Advanced project are welcome. Please refer to the contributing guidelines in the repository for more information.

## License

This project is licensed under the terms specified in the LICENSE file found in the root directory of the project.

## Sideloading the Add-in

To sideload the Writing Planner Advanced add-in in MS Word:

1. Clone or download the repository to your local machine.
2. Open MS Word and navigate to "Insert" > "My Add-ins" > "Shared Folder".
3. Select the folder containing the manifest.xml file from the repository.
4. The add-in should now be visible and available for use.

## Loading Organization Templates

To load organization-specific templates:

1. Ensure you have the necessary permissions and access to your organization's template repository.
2. Configure the template repository URL in the add-in settings.
3. Restart the add-in or reload MS Word to apply the changes.
4. Access the organization templates through the add-in's interface.