##Dynamics CRM Developer Extensions##

**Goal**

The goal of this project is to be a free alternative to the CRM Developer Toolkit that shipped with the Dynamics CRM 2011 & 2013 SDK. Currently it contains project  and item templates to help jump start the development process, code snippets, a tool to search for CRM related MSDN content, and a web resource deployer. The plan is to continue to expand and include other tooling to help streamline Dynamics CRM development. 

Supported versions of Visual Studio include 2012, 2013, & 2015 and will be distributed via the Visual Studio Gallery.

The project and item templates are based on the [SideWaffle](http://sidewaffle.com/) project.

**Installation**

Install in Visual Studio under Tools -> Extensions and Updates -> Search Online for "Dynamics CRM Developer Extensions" or install directly from the [Visual Studio Gallery](https://visualstudiogallery.msdn.microsoft.com/0f9ab063-acec-4c55-bd6c-5eb7c6cffec4).

####New in v1.1.0.0####

**Web Resource Deployer**

Right click and select the CRM Web Resource Deployer option to get started

* Manage mappings between multiple CRM organizations and Visual Studio projects
* Publish multiple items simultaneously from the interface or invidually by right clicking on the editor window or project item (must be mapped first)
* Filter by web resource type & managed/unmanaged
* Download web resources from CRM to your project
* Open CRM to view web resources
* Compare local version of mapped files with the CRM copy
* Add new web resources to CRM from a project file

**CRM SDK Search**

Select a block of text, right click and select CRM SDK to search MSDN filtering to Dynamics CRM content. 

**New User Options**

*Added option to allow publishing managed web resources     
*Added option to specify if the user's default browser or the Visual Studio browser is used for web content (SDK Search & open web resource in CRM)    
*Added option to disable SDK Search     

####v1.0.1.0####

**Templates**

Project Templates

* Plug-in   
* Plug-in Test   
* Custom Workflow Activity   
* Custom Workflow Activity Test   
* Web Resource   

Item Templates

* Plug-in Class   
* Plug-in Unit Test (MSTest)   
* Plug-in Integration Test (MSTest)   
* Plug-in Unit Test (NUnit)   
* Plug-in Integration Test (NUnit)   
* Custom Workflow Activity   
* Custom Workflow Activity Unit Test (MSTest)   
* Custom Workflow Activity Integration Test (MSTest)   
* Custom Workflow Activity Unit Test (NUnit)   
* Custom Workflow Activity Integration Test (NUnit)   
* JavaScript (Module) Web Resource   
* HTML Web Resource     

Here is an example of using the templates to create a plug-in project with unit tests:

[![Templates](http://img.youtube.com/vi/LdMhyF5x5jA/0.jpg)](https://youtu.be/LdMhyF5x5jA)

Review the [Wiki](https://github.com/jlattimer/CRMDeveloperExtensions/wiki) for additional documentation on using the templates.

**Code Snippets**

Currently there are JavaScript snippets for the majority of the 2011, 2013, and 2015 Dynamics CRM Client SDK. On the .NET side there are some snippets to add input and output parameters to custom workflow assemblies along with some snippets to assist in creating unit tests.

**Future**

Eventually there will be some functionality to help deploy custom assemblies from Visual Studio. I haven't thought this all the way through yet but currently I'm leaning toward still using the plug-in registration tool to do the initial deployment and creation of steps, images, etc... and then build something to update the assemblies when needed. The plug-in registration tool could be recreated inside Visual Studio but I'm not sure the effort to so would realistically save and significant amount of time.

If you have ideas for new templates or tools please post them in the issues area.

Feel free to [donate](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=KGV72FKEY8TJL) if this saved you some time or helped out :)
