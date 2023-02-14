# PDF Encrypter and Email Sender

This is a C# console application that encrypts PDF files with a password based on a national code and sends them to an email address specified in an Excel file.

## Prerequisites

To use this application, you will need:

- Visual Studio or another C# compiler
- .NET Framework 4.7 or later
- The iTextSharp and EPPlus packages, which can be installed using NuGet

## Getting Started

1. Clone or download the repository to your local machine.
2. Open the project in Visual Studio or another C# compiler.
3. Install the iTextSharp and EPPlus packages using NuGet.
4. Update the `pdfFolderPath`, `excelFilePath`, `emailSubject`, `emailBody`, `smtpServer`, `smtpPort`, `smtpUsername`, and `smtpPassword` variables in `Program.cs` with the appropriate values for your environment and use case.
5. Build and run the application.

## Functionality

When the application is run, it performs the following steps:

1. Reads in the national codes and email addresses from the specified Excel file.
2. Finds all PDF files in the specified directory.
3. Encrypts each PDF file with a password based on its national code.
4. Sends an email with the encrypted PDF file as an attachment to the specified email address.

If any errors occur during the execution of the application, they will be logged to both a text file and a Markdown file in the root directory of the project.

## Contributing

Contributions are welcome! If you find a bug or have an idea for a new feature, please create an issue or submit a pull request.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more information.
