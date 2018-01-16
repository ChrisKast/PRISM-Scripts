BlueZone Scripts | DHS-PRISM-Scripts
===

Table of contents for this README
---
*	**[Introduction](#introduction)**
*	**[State Administrators](#state-administrators)**
*	**[GitHub workflow and organization](#github-workflow-and-organization)**
*	**[Issue List](#issue-list)**
	*	[Issue guidelines/best practices](#issue-guidelinesbest-practices)
	*	[When to remove a script from the project](#when-to-remove-a-script-from-the-project)
* 	**[Pull Requests](#pull-requests)**
  	* 	[Anatomy of a pull request](#anatomy-of-a-pull-request)
    	* 	[Title](#title)
    	* 	["The Blip"](#the-blip)
	* 	[Pull Request Feedback](#pull-request-feedback)
  	* 	[When to wait on your pull request](#when-to-wait-on-your-pull-request)
  	* 	[What to do if there's a merge conflict](#what-to-do-if-theres-a-merge-conflict)
* 	**[Scriptwriters and Hackathons](#scriptwriters-and-hackathons)**
* 	**[About the Functions Library (FuncLib)](#about-the-functions-library-funclib)**
* 	**[About the Installer/Configuration Utility](#about-the-installerconfiguration-utility)**

---
Introduction
---

Welcome to the GitHub repository and project site for the PRISM BlueZone Scripts! This project (formally kicked off in July 2016) aims to automate repetetive, error prone tasks using simple extensions to the BlueZone Mainframe Display system. These scripts do not contain any confidential data, nor do they contain information about how to log in to our various state systems.

If you have questions about BlueZone Scripts and work in a Minnesota human services agency, please ask a supervisor about getting started.

_[(back to top)](#bluezone-scripts--dhs-prism-scripts)_

---
State Administrators
---
As of January 2017, there is one statewide coordinator for the BlueZone Scripts project:
*	**For PRISM scripts: [Veronica Cary](mailto:veronica.cary@state.mn.us)**
*	**For MAXIS scripts:** vacant, Veronica Cary filling in as needed

In addition to our statewide coordinator, numerous county/tribal agency personnel are involved in much of the writing and testing of BlueZone Scripts. Their contribution is appreciated!


_[(back to top)](#bluezone-scripts--dhs-prism-scripts)_

---
GitHub workflow and organization
---

GitHub can be somewhat complex and daunting for beginners. For our organization, scripts are divided into two "branches":

*	**master**: the working directory for scriptwriters and select power users in scriptwriting agencies.
* 	**release**: the branch for most eligibility workers statewide.

Changes proposed here are evaluated, cataloged, and acted on using a standard workflow:
1.	A bug is reported, or a new script idea is suggested, using an issue on our GitHub issues list.
2.	The statewide coordinator reviews the issue for basic technical feasibility and possible redundancy (copies of other issues or 
	scripts which already exist).
	*	If the suggestion is a brand new script, or an enhancement to a current script which may impact policy, pertinent staff from DHS 
		policy areas review the suggestion. Script ideas that do not meet their approval will be rejected with the reason posted on the 
		GitHub issue.
3.	The suggested idea is posted for any possible scriptwriters from counties who want to take it. As the scriptwriters work for their 
	counties, and not DHS, there is no guarantee that new ideas will be acted on. However, every attempt will be made to ensure script 
	ideas are implemented as quickly as possible. Issues not currently "claimed" by a scriptwriter will be tagged as "backlogged" in our 
	issue list. In some cases, the statewide coordinator may work on an issue directly, particularly if it fixes a bug or provides a 
	substantive benefit to staff (these will be evaluated on a case-by-case basis).
4.	If the issue is taken up by a scriptwriter (county or state), it will be processed as soon as possible, and the statewide 
	coordinator will check in with the scriptwriter throughout the process of writing the script.
5.	Once the script is developed by the scriptwriter, they are expected to test the new script locally and collect feedback from any 
	pertinent local staff involved. After that point, the scriptwriter submits a *pull request* containing the script, which is reviewed 
	by the statewide coordinator before being pulled into the master branch for testing. **Only the statewide coordinator can pull in 
	changes into the supported branches.**
6.	Agencies that are using the master branch (the branch used for testing) are notified of the new/improved script, and are asked to 
	test it and report back to the statewide coordinator.
7.	Once testing is determined to be successful, the new/improved script is pulled into the release branch.

_[(back to top)](#bluezone-scripts--dhs-prism-scripts)_

---
Issue list
---

We have an [issue list](https://github.com/MN-Script-Team/DHS-PRISM-Scripts/issues) maintained on GitHub. Both scriptwriters and non-scriptwriters should feel free to create/report issues on the issue list (doing so requires a GitHub account).

Scriptwriters are encouraged to tackle any issues on the issue list, so long as they meet the following conditions:
* The scriptwriter has time in the near future to complete and test the issue (note that many issues have a "milestone", which may have an associated due date).
* The scriptwriter adds a comment to the issue saying they'll take it.
* The scriptwriter builds (or modifies) instructions on SIR after completing the work.
* For new scripts, the scriptwriter tests the new script on multiple cases/scenarios before submitting (ideally, for a week or so on active cases in their agency).

**ABSOLUTELY NO CLIENT DATA SHOULD EVER BE SHARED ON GITHUB.** In addition, please refrain from posting entire screenshots of system screens on GitHub issues. If case numbers or screenshots are needed, please share them via secure email (see your agency for your local process).

#### Issue guidelines/best practices
* Search existing issues before submitting a new one. Duplicates are annoying and add unneccessary work for administrators (as well as duplicate email notifications). It may also be wise to search through closed issues (by selecting "closed" in the top of the issue list).
* Issue title should be short (under 75 characters, or about the size of a case note header). This goes in the subject line for emails, so keep it clean.
* For existing scripts, please indicate the script category/name at the beginning of the issue (ex. "NOTES - CAF: needs longer space for 'other notes'"). This is helpful for organization.
* If there are multiple issues with an existing script, create separate issues for each. This is easier both for release notes tracking and for recipients of GitHub update emails.
* Don't upload screenshots of code, as it does not meet accessibility standards (and can't easily be copy/pasted). If you want to discuss code snippets, copy/paste them and surround them in blocks using GitHub markdown's default format (3 backticks).
* If you have a question, it should only be posted if you believe a change to a script is necessary or wise. If it's a general scripts question, it is better addressed via email.

#### When to remove a script from the project
Some scripts will be very popular in some agencies, and not-so-popular in others; such is the nature of a large collaborative project. It may be tempting to advocate for the removal of a script based on preferences in your agency, or even preferences among scriptwriter agencies as a whole. However, each script represents someone's hard work, and should be treated with respect. Generally speaking, this project will strive to improve existing scripts, and not to remove scripts from the project. Even so, there are going to be circumstances in which removing a script could become necessary:

* An older script is completely superseded by another, newer script. For example, a NOTES script becoming an ACTIONS script.
* A policy change (or clarification) has made the script completely unnecessary, or even incorrect.

In all other instances, the work should be toward the improvement and augmentation of our scripts. Discussions about removing a script must remain respectful of the work other scriptwriters have done, and the script must meet either of the criteria above in order to be considered for removal.

_[(back to top)](#bluezone-scripts--dhs-prism-scripts)_

---
Pull Requests
---
A "pull request" is the process of requesting that the script administrator "pull" your changes in to the main branch. The process of making a pull request is fully documented elsewhere.

#### Anatomy of a pull request

Pull requests, when done correctly, can make changes easy for script administrators.

##### Title
A proper pull request first contains the issue it resolves. Immediately following this is a short (30-50 characters) explanation of what script this relates to, and what your change contains.

For example:

> `#24601: readme contains correct link`

If your pull request encompasses multiple issues, list them individually:

> `#24601, #24602, #24603`

**NOTE**: your pull request should typically only address a single script, unless a single issue spans multiple scripts.

Pull requests that deviate from this process might be rejected. If you make a mistake, you can always update the title of a pull request after you send it. This is an important process in order to maintain the integrity of our release notes and email/RSS notifications.

##### Updating the changelog
It is important to summarize your update in a short "blip" on the changelog within each script. _This should be written from a non-technical perspective_ that end users would understand. The changelog syntax is:

> `call changelog_update("<<DATE>>", "<<DESCRIPTION>>", "<<SCRIPTWRITER>>")`

A good changelog example:

> `call changelog_update("01/01/2000", "An issue with sending DORD Z1234 has been resolved.", "Jane Public, Doe County")`

A bad changelog example:

> `call changelog_update("01/01/2000", "A function was broken. Replaced function with new", "Jane")`

Note in the latter a lack of detail from an end user perspective (what did the new function fix?), as well as an incomplete sentence structure. Also note an incomplete name for the scriptwriter. To support the clear language goals of the department, the changelog should be simple, clear, and concise.

If a changelog is missing or incomplete, your pull request might be rejected. Remember, you can always update an existing pull request if needed.

#### Pull Request Feedback
This is a collaborative project, and feedback on your pull requests is bound to come in from state and county/agency staff.

* Those who give feedback are expected to be considerate and respectful of the scriptwriter and their work.
* Scriptwriters receiving feedback are expected to incorporate suggestions or explain any disagreements/concerns (in a respectful manner).

We are working with people's creative output here. Disrespectful comments or unproductive suggestions will not be tolerated. But, suggestions to improve code readability, functionality, or consistency are expected to be followed (particularly coming from state administrators). Discussion is always welcome, so long as it's respectful.

#### What to do if there's a merge conflict
A "merge conflict" is a conflict in which one version of a file conflicts with another change. Perhaps you made a change at the same time someone else did? That's usually a pretty common reason for the conflict. Typically conflicts can be resolved in a few different ways:
* Compare each change against the most recent change to the file in the master branch. Sometimes it's really easy to tell where the conflict is when you look at the most recent change someone else made. If you suspect that's the issue, copy the new version into your change (within your branch), and see if that solves the problem.
* Try doing a pull request from the master branch to your branch (like a "reverse" pull request). This sometimes fixes issues.
* If you're using GitHub Desktop, try following the recommended "command line instructions" (which are included within the pull request).
* If all else fails, contact a state administrator and ask them! We're ready to help resolve these (but it might take a bit of time).

_[(back to top)](#bluezone-scripts--dhs-prism-scripts)_

---
Scriptwriters and Hackathons
---
A typical (potential) BlueZone Scriptwriter has excellent critical thinking skills, is well organized, and reasonably computer savvy. Previous coding experience isn't required, but definitely helps. We use VBScript predominantly, run through the BlueZone Script Host (taking advantage of the custom functions that we've built, and the built-in custom functions of BlueZone). There are numerous online courses for VBScript and computer programming in general.

Trainings for BlueZone Scriptwriters are conducted on occasion. Contact a script administrator if you're interested in this.

BlueZone Scripts "hackathons" take place the 2nd and 4th Monday of each month. Any trained scriptwriter is permitted to attend with permission from their supervisor and state administrators. Contact a script administrator to find out more.

Our team maintains a [Slack organization](https://slack.com) which county and state scriptwriters can use to communicate quickly and easily. The organization can be found at [mn-script-team.slack.com](https://mn-script-team.slack.com).


_[(back to top)](#bluezone-scripts--dhs-prism-scripts)_

---
About the Functions Library (FuncLib)
---
The BlueZone Scripts "Functions Library" (or "FuncLib") is a centrally located repository for the common "shared functions" associated with MAXIS and PRISM scripts. We use a separate repository for it, as it does not always follow the same release cycle as other scripts, and we want to keep the highly-technical conversation (and associated notifications) to a minimum for non-technical followers of our work. The FuncLib repository can be [found here](https://github.com/MN-Script-Team/BZS-FuncLib).

The "FuncLib block" is a block of code used within scripts that incorporates the shared Functions Library- this allows scriptwriters to take advantage of the dozens of functions used for navigation, case noting, autofill, and more. To use this library, simply add this block of code to the top of your script:
``` vbnet
'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================
```

_[(back to top)](#bluezone-scripts--dhs-prism-scripts)_

---
About the Installer/Configuration Utility
---
The BlueZone Scripts are "installed" (more like "configured" as they are simply script files inserted into a directory) using a utility [located here](https://mn-script-team.github.io/bzs-config-utility/).
