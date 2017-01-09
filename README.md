##Dynamics CRM Developer Extensions##

**Goal**

The goal of this project is to be a free alternative to the CRM Developer Toolkit that shipped with the Dynamics CRM 2011 & 2013 SDK. Currently it contains project and item templates to help jump start the development process, code snippets, a tool to search for CRM related MSDN content, and tooling to assist with managing and deploying web resources, plug-ins, and reports. The plan is to continue to expand and include other tooling to help streamline Dynamics CRM development. 

Supported versions of Visual Studio include 2012, 2013, & 2015 and will be distributed via the Visual Studio Gallery.

**Installation**

Install in Visual Studio under Tools -> Extensions and Updates -> Search Online for "Dynamics CRM Developer Extensions" or install directly from the [Visual Studio Gallery](https://visualstudiogallery.msdn.microsoft.com/0f9ab063-acec-4c55-bd6c-5eb7c6cffec4).

####New in v1.3.4.1####
Updated connection string for on premise w/ credentials  
Added missing quickform snippets

####New in v1.3.4.0####

* CRM/D365 v8.2 support
* Updated to handle 8.2 SDK assemblies
* Added 8.2 JavaScript snippets
* Changed JavaScript snippet format to be easier to use - prefixes start with "Xrm"
* **Changed CRM connection to use Xrm.Tooling - this is a breaking change!**
   * For existing projects you will either need to:
      * Create a new connection & remap items
      * Open the CRMDeveloperExtensions.config at the project root and replace the Base64 encoded connection string with one created from a new connection or a Base64 encoded version from: [https://msdn.microsoft.com/en-us/library/mt608573.aspx] (https://msdn.microsoft.com/en-us/library/mt608573.aspx)
* Added OAuth as a new connection type
* Added option to enable verbose logging from Xrm.Tooling
* Updated project templates to use Xrm.Tooling for connections with SDK versions 8.0+
* Various bug fixes and usability enhancements

**Solution Packager**

* 1 click download and extraction of CRM solution to a Visual Studio project
* Re-package solution files from project

Review the [Wiki](https://github.com/jlattimer/CRMDeveloperExtensions/wiki/Solution-Packager) for additional documentation. 

**Plug-in Deployer**

* 1 click deploy plug-ins and custom workflow assemblies from Visual Studio without having to click through the SDK plug-in registration tool
* Integrated ILMerge on plug-in and custom workflow projects

Review the [Wiki](https://github.com/jlattimer/CRMDeveloperExtensions/wiki/Plugin-Deployer) for additional documentation. 

**Report Deployer**

* 1 click deploy reports from Visual Studio without having to go through CRM
* Clear local dataset cache 

Review the [Wiki](https://github.com/jlattimer/CRMDeveloperExtensions/wiki/Report-Deployer) for additional documentation.

**Web Resource Deployer**

* 1 click deploy of web resources from Visual Studio without having to go through CRM
* Download web resources from CRM to Visual Studio project
* Diff local files with CRM server versions

Review the [Wiki](https://github.com/jlattimer/CRMDeveloperExtensions/wiki/Web-Resource-Deployer) for additional documentation. 

**Templates**

Project Templates

* Plug-in   
* Plug-in Test   
* Custom Workflow Activity   
* Custom Workflow Activity Test   
* Web Resource   
* TypeScript

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

**Code Snippets**

* JavaScript code snippets for CRM 2011, 2013, 2015, & 2016


If you have ideas for new templates or tools please post them in the issues area.

Feel free to [donate](https://www.paypal.me/JLattimer) if this saved you some time or helped out :)
