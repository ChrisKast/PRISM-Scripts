# BlueZone Scripts Contribution Guidelines

The BlueZone Scripts PRISM project welcomes any issues, enhancement ideas, or pull requests for improving our work. The following are our very brief general guidelines, and more can be seen in our README.md file.

## For creating issues
- Title your issue with the name of the script and a brief description, such as "Intake script: DORD process is broken".
- Where possible, use the term "bug" to refer to a *problem*, and "enhancement" to refer to a general improvement. Calling a bug an "enhancement request" might lead to a miscalculation of the time needed to get the work done.
- **DO NOT POST SYSTEM SCREENSHOTS OR CLIENT DATA ANYWHERE ON ISSUES**. If you need to submit them, do so via secure email per standard county/state processes.

## For creating pull requests
- Ensure your file is in the right folder. For example, any actions scripts should be in the "actions" folder.
- Ensure your file also follows the proper formatting (all lower case, no spaces). So, a script called "Sends Z1234 DORD" would be located at `actions/sends-z1234-dord.vbs`.
- Title your pull request referencing the issue number, and a brief description of what you did, such as "#12345: fixed DORD process in Intake script"
- Update the changelog in the script. The changelog syntax is `call changelog_update("01/01/2000", "An issue with sending DORD Z1234 has been resolved.", "Jane Public, Doe County")`. Ensure that the syntax works by testing your script *with* the changelog syntax updated. For the description, remember that it will be read by non-scriptwriters, so don't get too technical.
- Please keep in mind that all pull requests will be reviewed. Feedback will be sent, and you will be expected to address it ASAP. Pull requests that have been waiting for longer than a few days without resolution may be closed.
- All submissions will be reviewed by DHS staff. Policy-specific questions may escalate to other staff in the agency.
