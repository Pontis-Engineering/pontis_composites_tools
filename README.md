# Pontis Composites Custom Tool for Femap

This repository contains a collection of scripts that form a Custom Tool you can easily add directly into Femap. This is done using Femap API and is geared around creating FE models for composite materials.

These scripts are designed to enhance your Femap workflow, automate repetitive tasks, and extend the capabilities of the software. Some are generic, however most are focused around pre/post processing FEM models representing structures build using composite materials.

We only encourage you to create a GitHub account (free to do) so you can leave feedback and contribute.

# Background
At Pontis we often use Femap in our design process. Femap is just one of several Finite Element packages you can use for structural analysis. In general, it is just more complex to create models and optimise designs in composite. To help here more specific and advanced software tools can be used to achieve more efficient and effective workflows. These composite tools can be either commercially available, developed in-house, or a combination of both.  In this regard, over time at Pontis we have developed our own tools that compliment existing Femap capability.
 
We often get asked about these tools and so in the spirt of enabling people to work more easily in Femap with composites we want to share a free to use custom tool add-on for Femap. Initially they are just a few functions (e.g. importing/exporting layups) but we believe still very useful to speed up your workflow and introduces you to the possibilities. If there is enough interest, we plan to continue to add more functions and upgrades 😉.

# Current Functions 
- Create/Extract Materials.
- Create/Extract Properties [e.g. laminates].
- Create/Extract Layups [i.e. ply tables]
# Instructions!
- The extact_to_file functions will open and excel file and worksheet (e.g. materials).
- The create_from_file require to select an excel file with relevent worksheet (e.g. materials).
- Hint: The format of this worksheet can be determined by first using the extact_to_file function.
- Note: Density unit is assumed to be in kg/m^3 but converted to kg/mm^3 (i.e. assuming your model is mm).
- Tip: The column heading name should not change, but the order you can change, for example you could add a column E33 if you also want to extract that value.
- Hint: We intend to keep improving the functionality and resilient, so please raise any issues you find.

# Getting Started [Installing the Custom Tool]

## Prerequisites
- Femap installed on your system.
- Fundamental understanding of Femap's user interface and scripting capabilities.

## Installation
There are two possible options to enable custom scripts in your Femap. Below both are explained, choose one.

### 1. Default Femap Location

- Clone or download this repository to your local machine.
- Go to directory where Femap is installed and find **api** folder. 
- Default location of the folder: `C:/Program Files/Siemens/Femap [VERSION NO.]/api`
- Paste the contents of this repository to this folder. 
- Restarting the Femap might be necessary.

### 2. Set custom scripts location

*This is preferred option if you cannot paste files into your `Program Files`*

- Clone or download this repository to your local machine.
- Move downloaded files to a prefered location on your hard drive.
- Open femap and go to User Tools (see image below).

![User Tools.png](assets%2FUser%20Tools.png)

- when you click Tools Directory, you will be prompted with the file explorer. 
- Select the folder where you store your scripts and confirm the operation.
- Scripts should be available under `User Tools` menu.

## Usage
- Open Femap.
- Depending on the installation method go either to `Custom Tools` or `User Tools`.
- Adhere to any instructions on the screen when prompted.

## License
Code in this repository is shared free of charge, licensed under the Apache 2.0 License.
For bespoke inquires please contact us via <br>
info\@pontis-engineering.com
