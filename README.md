# Daily Rotation Bot
A MS Teams bot that helps you automate rotations of people. 

An example of how it can be used:     

![image](https://user-images.githubusercontent.com/9387558/207445442-4d9ff792-80e1-46e9-8fe3-f0d243eb96c4.png)

## Prerequisites

- [NodeJS](https://nodejs.org/en/)
- [Teams Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) or [TeamsFx CLI](https://aka.ms/teamsfx-cli)

## Debug

- From Visual Studio Code: Start debugging the project by hitting the `F5` key in Visual Studio Code. 
- Alternatively use the `Run and Debug Activity Panel` in Visual Studio Code and click the `Run and Debug` green arrow button.
- From TeamsFx CLI: Start debugging the project by executing the command `teamsfx preview --local` in your project directory.

## Edit the manifest

You can find the Teams app manifest in `templates/appPackage` folder. The folder contains one manifest file:
* `manifest.template.json`: Manifest file for Teams app running locally or running remotely (After deployed to Azure).
