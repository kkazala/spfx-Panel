# Panel for ListView Command Set extensions

## Summary

This control renders stateful Panel that can be used with [ListView Command Set extensions](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/building-simple-cmdset-with-dialog-api). It may optionally refresh the list view page after the panel is closed.
It opens when a List Command button is clicked, and closes using either Panel's close button, or on "light dismiss".

It may be used to replace Dialog component, ensuring the User Interface is consistent with that of SharePoint Online.

![picture of the extension in action](https://github.com/kkazala/spfx-Panel/blob/main/assets/PanelSpfx1.14.gif)

## Compatibility

![Compatible with SharePoint Online](https://img.shields.io/badge/SharePoint%20Online-Compatible-green.svg)
[![version](https://img.shields.io/badge/SPFx-1.16.1-green)](.1) ![version](https://img.shields.io/badge/Node.js-16.18-green)
![Hosted Workbench Compatible](https://img.shields.io/badge/Hosted%20Workbench-Compatible-green.svg)

![pnp V3](https://img.shields.io/badge/pnp-V3-green)
![TypeScript-4.5](https://img.shields.io/badge/TypeScript-4.5-green)
![rush--stack--compiler-4.5](https://img.shields.io/badge/%40microsoft%2Frush--stack--compiler-4.5-green)

![Does not work with SharePoint 2019](https://img.shields.io/badge/SharePoint%20Server%202019-Incompatible-red.svg "SharePoint Server 2019 requires SPFx 1.4.1 or lower")
![Does not work with SharePoint 2016 (Feature Pack 2)](<https://img.shields.io/badge/SharePoint%20Server%202016%20(Feature%20Pack%202)-Incompatible-red.svg> "SharePoint Server 2016 Feature Pack 2 requires SPFx 1.1")
![Local Workbench Incompatible](https://img.shields.io/badge/Local%20Workbench-Incompatible-red.svg)

> As of [SPFx 1.16](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/release-1.16):
>
> -   Node.js v12 & v14 are no longer supported. **SPFx v1.16 requires Node.js v16.**
> -   `@microsoft/office-ui-fabric-react-bundle` is DEPRECATED. Use `@fluentui/react` instead.

## Applies to

-   [SharePoint Framework](https://aka.ms/spfx)
-   [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

SPFx development environment compatibility:
| SPFx | Node.js (LTS)| NPM| TypeScript| React| @fluentui/react |
|-|-|-|-|-|-|
|1.16.1| v16.13+| v5, v6, v7, v8| v4.5| v17.0.1| v8|

> Important:
> In order to use [React plug-in for Application Insights JavaScript SDK](https://learn.microsoft.com/en-us/azure/azure-monitor/app/javascript-react-plugin) components, set `"allowSyntheticDefaultImports": true` in `tsconfig.json`.
>
> Use `@fluentui/react` **v8** for compatibility with **React v17**

## Solution

| Solution   | Author(s)    |
| ---------- | ------------ |
| spfx-panel | Kinga Kazala |

## Version history

| Version | Date              | Comments                                                                      |
| ------- | ----------------- | ----------------------------------------------------------------------------- |
| 1.3     | February 02, 2023 | Correct version of @fluentui/react                                            |
| 1.2     | January 20, 2023  | Upgrade to SPFx 1.16 which requires Node.js v16. Application Insights support |
| 1.1     | March 28, 2022    | Upgrade to SPFx 1.14, Update item using pnp V3                                |
| 1.0     | January 13, 2022  | Initial release                                                               |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

-   Clone this repository
-   Ensure that you are at the solution folder
-   in the command-line run:
    -   **nvm use 16.18.x** (the 16.x.x node version you have installed)
    -   **npm install**
    -   **gulp serve --nobrowser** (or **gulp serve --nobrowser --locale=de-de** to test German version)
    -   debug

See [Debugging SPFx 1.13+ solutions](https://dev.to/kkazala/debugging-spfx-113-solutions-11cd) on creating debug configurations.

## Features

Opening and closing Panel controls is a no-brainer as long as it is controlled by a parent component.
In the case of a ListView Command Set, this requires slightly more effort.

This extension illustrates the following concepts:

-   Panel component with an [AppInsightsErrorBoundary](https://learn.microsoft.com/en-us/azure/azure-monitor/app/javascript-react-plugin#react-error-boundaries)
-   Logging using @pnp/logging Logger, configured to use ConsoleListener and (if application insights connection string is provided), custom AppInsightsLogListener
-   page tracking and event tracking with Application Insights
-   Example component using Panel, with a Toggle control to refresh the page when the panel is closed

### React Error Boundary

As of React 16, it is recommended to use error boundaries for handling errors in the component tree.
Error boundaries **do not catch** errors for event handlers, asynchronous code, server side rendering and errors thrown in the error boundary itself; try/catch is still required in these cases.
This solution uses [AppInsightsErrorBoundary](https://learn.microsoft.com/en-us/azure/azure-monitor/app/javascript-react-plugin#react-error-boundaries) component.

### Avoiding "The security validation for this page is invalid and might be corrupted"

"The security validation for this page is invalid and might be corrupted" issue only occurs when the spfi() object and spfx behavior is instantiated multiple times during the lifecycle of a web part.
See: https://github.com/pnp/pnpjs/issues/2304

### PnP Logger

Logging is implemented using [@pnp/logging](https://pnp.github.io/pnpjs/logging) module. [Log level](https://pnp.github.io/pnpjs/logging/#log-levels) is defined as an extension property, which allows changing log level of productively deployed solution, in case troubleshooting is required.

Errors returned by [@pnp/sp](https://pnp.github.io/pnpjs/sp/#pnpsp) commands are handled using `Logger.error(e)`, which parses and logs the error message. If the error message should be displayed in the UI, use the [handleError](src\extensions\utils\ErrorHandler.ts) function implemented based on [Reading the Response](https://pnp.github.io/pnpjs/concepts/error-handling/#reading-the-response) example.

### Application Insights

This solution is using [Application Insights for webpages](https://learn.microsoft.com/en-us/azure/azure-monitor/app/javascript) and [React plug-in for Application Insights JavaScript SDK](https://learn.microsoft.com/en-us/azure/azure-monitor/app/javascript-react-plugin) to log errors and metrics to application insights.

## Deploy

Install solution

```powershell
# -Scope accepted values: Tenant, Site
$packageInSite = Add-PnPApp -Path "$sppkgPath" -Scope $appCatalogScope -Overwrite -Publish
if ( $null -eq $packageInSite.InstalledVersion ) {
    Write-Host "Installing app $($packageInSite.Id) ..."
    Install-PnPApp -Identity $packageInSite.Id -Scope $appCatalogScope -Wait
}
elseif ($packageInSite.CanUpgrade -eq $true) {
    Write-Host "Updating installed app $($packageInSite.Id) ..."
    Update-PnPApp -Identity $packageInSite.Id -Scope $appCatalogScope
}
```

In case you are not using the elements.xml file for deployment, you may add the custom action using `Add-PnPCustomAction`

```powershell
Add-PnPCustomAction -Title "Panel" -Name "panel" -Location "ClientSideExtension.ListViewCommandSet.CommandBar" -ClientSideComponentId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ClientSideComponentProperties "{""listName"":""Travel Requests"",""logLevel"":""3"",""appInsightsConnString"":""your-connection-string""}" -RegistrationId 100 -RegistrationType List -Scope Web
```

Updating the properties in an already deployed solution can be done with:

```powershell
$ca=Get-PnPCustomAction -Scope Web -Identity "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
$ca.ClientSideComponentProperties="{""listName"":""Travel Requests"", ""logLevel"":""1"",""appInsightsConnString"":""your-connection-string""}"
$ca.Update()
```

## References

-   [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
-   [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
-   [Debugging SPFx 1.13+ solutions](https://dev.to/kkazala/debugging-spfx-113-solutions-11cd)
-   [Professional SPFx Solutions: Superb solution packages](https://pnp.github.io/blog/post/spfx-21-professional-solutions-superb-solution-packages/)
-   [PnP Error Handling](https://pnp.github.io/pnpjs/concepts/error-handling/)
-   [PnP log levels](https://pnp.github.io/pnpjs/logging/#log-levels)
-   [https://pnp.github.io/pnpjs/logging/#create-a-custom-listener](https://pnp.github.io/pnpjs/logging/#create-a-custom-listener)
-   [React plug-in for Application Insights JavaScript SDK](https://learn.microsoft.com/en-us/azure/azure-monitor/app/javascript-react-plugin)
-   [React Error Boundaries](https://reactjs.org/docs/error-boundaries.html) in React 16, and [React error boundaries](https://learn.microsoft.com/en-us/azure/azure-monitor/app/javascript-react-plugin#react-error-boundaries)
-   [I Made a Tool to Generate Images Using Office UI Fabric Icons](https://joshmccarty.com/made-tool-generate-images-using-office-ui-fabric-icons/) to generate CommandSet icons

### React

> It is [recommended](https://beta.reactjs.org/reference/react/PureComponent#migrating-from-a-purecomponent-class-component-to-a-function) to use function components instead of class components in the new code.

-   [Understanding Functional Components vs. Class Components in React](https://www.twilio.com/blog/react-choose-functional-components)
-   [Components and Props](https://reactjs.org/docs/components-and-props.html)
-   [Understanding the importance of the key prop in React](https://dev.to/francodalessio/understanding-the-importance-of-the-key-prop-in-react-3ag7=)
-   [Understanding React's key prop](https://kentcdodds.com/blog/understanding-reacts-key-prop)
