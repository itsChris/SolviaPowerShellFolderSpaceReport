# Folder Size Report Generator

This PowerShell script calculates the sizes of folders starting from a specified entry point and generates an HTML report including a summary of the total size and detailed information for each folder. The report features hierarchical presentation of folders, their paths, sizes, and includes a company logo for a professional appearance.

## Features

- **Recursively calculates folder sizes:** Specify a root folder and depth for analysis.
- **Generates an HTML report:** Creates a visually appealing report with CSS styling.
- **Includes a summary of the total size:** A summary of the total size of the root folder is displayed at the top of the report.
- **Customizable:** Easily adaptable script to include more details or modify the report layout.

If you're seeking a TreeSize alternative, this PowerShell script serves as a powerful and customizable solution. Unlike TreeSize, which is a graphical application for managing disk space, this script provides a more flexible, script-based approach. It allows for automated folder size calculations and report generation, tailored specifically to your needs. Here are some key points that make this PowerShell script a noteworthy alternative:

- **Customization**: Unlike TreeSize, every aspect of this script can be customized to fit specific requirements, from the depth of folder analysis to the appearance of the report.
- **Automation**: The script can be incorporated into larger automation tasks, making it a versatile tool for system administrators.
- **No Installation Required**: Unlike TreeSize, this script doesn't require installation of additional software if you already have PowerShell, making it a lightweight alternative.
- **Free and Open Source**: Being a script, it's inherently open-source and free to use, modify, and distribute, offering a cost-effective alternative to TreeSize.

This PowerShell script is ideal for users who prefer a programmable approach to disk space management, offering a blend of flexibility, power, and simplicity.

## Getting Started

### Prerequisites

- Windows PowerShell 5.1 or PowerShell 7.x
- Basic understanding of PowerShell scripting

### Installation

1. Clone the repository or download the script directly.
2. Ensure the `assets` folder containing the `solvia.svg` logo is in the same directory as the script or adjust the path in the script accordingly.

### Usage

To run the script, open PowerShell and navigate to the script's directory. Execute the script with the following command, replacing the placeholders with your desired values:

```powershell
$folderData = DetermineFolderSizes -FolderEntryPoint "<YourFolderPath>" -Depth <Depth>
CreateReport -FolderData $folderData
```

## Customization
You can modify the PowerShell script and the HTML/CSS for branding or preferences.

## Contributing
Contributions are welcome. Fork the repository, make your changes, and submit a pull request.

## Author
Christian Casutt - Solvia GmbH

## License
This project is licensed under the MIT License.

## Acknowledgments
Thanks to Solvia GmbH and the PowerShell community.