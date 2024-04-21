# Femap Automation Repository

This repository contains a collection of scripts and utilities for use with Femap, a powerful finite element modeling software. 
These scripts are designed to enhance your Femap workflow, automate repetitive tasks, and extend the capabilities of the software.
Some of those will be generic, however most of them will be focused around pre/post processing FEM models representing structures build using composite materials.

## Contents
- This section will be gradually updated once scripts will be ready to be published. 
## Getting Started

### Prerequisites
To ensure a smooth process, verify that you meet the following prerequisites:

- Femap installed on your system.
- Fundamental understanding of Femap's user interface and scripting capabilities.

### Installation
There are two possible options to enable custom scripts in your Femap. Below both are explained, choose one.

#### Default Femap Location

- Clone or download this repository to your local machine.
- Go to directory where Femap is installed and find **api** folder. 
- Default location of the folder: `C:/Program Files/Siemens/Femap [VERSION NO.]/api`
- Paste the contents of this repository to this folder. 
- Restarting the Femap might be necessary.

#### Set custom scripts location

*This is preferred option if you cannot paste files into your `Program Files`*

- Clone or download this repository to your local machine.
- Move downloaded files to a prefered location on your hard drive.
- Open femap and go to User Tools (see image below).

![User Tools.png](assets%2FUser%20Tools.png)

- when you click Tools Directory, you will be prompted with the file explorer. 
- Select the folder where you store your scripts and confirm the operation.
- Scripts should be available under `User Tools` menu.

### Usage
- Open Femap.
- Depending on the installation method go either to `Custom Tools` or `User Tools`.
- Adhere to any instructions on the screen when prompted.

### License
Code in this repository is shared free of charge, licensed under the Apache 2.0 License.
For bespoke inquires please contact us via [info\@pontis-engineering.com](mailto:info@pontis-engineering.com?subject=Automation inquiry)