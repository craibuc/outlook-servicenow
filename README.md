# outlook-servicenow

Service Now tools for Outlook

# Features

* Uses the RITM # in a message's subject or body to messages organized, placing the message in an eponymously-named folder

# Installation

* Open Outlook
* Press <kbd>ALT</kbd>+<kbd>F11</kbd>
* In the Project browser, open the Modules node, select `Import file...`, then choose `Rules.bas`

# Usage

* Create a new Outlook rule (`Rules | Create Rule...` menu)
* Select the <kbd>Advanced Options...</kbd> Button
* Choose desired `Select condition(s)`, then click <kbd>Next</kbd>
* Select `Run a Script`, then click `script` in the Edit rule window
* Select `Project1.ProcessMailItem` from the `Select Script` window

If the `Run a Script` option is not present, please follow these instructions to enable it: https://www.extendoffice.com/documents/outlook/4640-outlook-rule-run-a-script-missing.html.

# Dependencies

Located in project's `Tools | References...` window:
* Microsoft Scripting Runtime
* Microsoft VBScript Regular Expressions 5.5

# Contributions

* Import the `RulesTest.bas` file; ensure that unit test pass
