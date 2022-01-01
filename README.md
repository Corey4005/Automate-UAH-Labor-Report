# A Simple Shell Script to Automate Filling in the UAH Labor Report
This is a simple shell script to make filling out the UAH labor report easier. The script takes in data from a template [text file](./edit-this-text-file.txt). Then, it uses this template to fill in associated cells in this year's 2022 labor report found [here](./excel-data/Labor-Report-2022.xlsx)

## Installation
First, open a terminal window in [*Anaconda Prompt*](https://docs.anaconda.com/anaconda/install/windows/) and cd to a directory where you would like the files from this repository to be stored. Then, clone the repository using the following command:

```bash
$ git clone https://github.com/Corey4005/Automate-UAH-Labor-Report.git
```
After, install the packages needed for this project by using the [requirements.txt](./requirements.txt) file in the [Automate-UAH-Labor-Report](./) directory. 

**NOTE**: if you do not want to install the libraries for this project on your root folder,  you may want to first create a virtual environment with the following command (where *yourenvname* is a variable for what you would like to call it):

To Create a Virtual Environment:
```
$ conda create -n yourenvname
```
To Install the Necessrry Packages:

```
$ pip install -r requirements.txt
```
## Usage

Open the [text file](./edit-this-text-file.txt) in your local directory and replace the template with your information: 

### Template *Before Edits:*

![Template Before Edits](./images-for-read-me/Template.png)

### Template *After Edits*
![Template Edits](./images-for-read-me/Template-Edit.png)

Save the template edits and navigate to the script file in your terminal. After run the script and input the information that is prompted to you: 

```
$ python script.py
```

Re-open the now edited excel file and viola! **You now have automated inputing info into your UAH labor report!**

# Would you like to contribute? 

Here are some things that may be helpful: 

Task-list: 
- [x] Automated Inputs for UAH Labor Report
- [ ] Work on Automated Email that sends attached Excel Work from user to boss via smtp server. 
- [ ] Create functionality that will allow user to write multiple sheets at a time with one command. 