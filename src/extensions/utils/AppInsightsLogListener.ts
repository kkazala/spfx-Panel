
import { ReactPlugin } from '@microsoft/applicationinsights-react-js';
import { ApplicationInsights, SeverityLevel } from '@microsoft/applicationinsights-web';
import { ILogEntry, ILogListener, LogLevel } from "@pnp/logging";
import { createBrowserHistory } from "history";
// https://www.sharepointpals.com/post/how-to-log-spfx-react-webpart-using-azure-application-insights-and-pnp-logging/

export interface IAppInsightsLogListenerProps{
    connectionString:string;
    version: string; //this.manifest.version
}

//https://learn.microsoft.com/en-us/azure/azure-monitor/app/javascript-react-plugin#basic-usage
export class AppInsightsLogListener implements ILogListener {
    private static appInsights : ApplicationInsights;
    private static reactPlugin : ReactPlugin;

    constructor(props: IAppInsightsLogListenerProps) {
        if (!AppInsightsLogListener.appInsights )
            AppInsightsLogListener.appInsights = AppInsightsLogListener.initialize(props);
    }

    private static initialize(props?: IAppInsightsLogListenerProps): ApplicationInsights {
        const browserHistory = createBrowserHistory({  });
        AppInsightsLogListener.reactPlugin  = new ReactPlugin();
        const appInsights = new ApplicationInsights({
            // https://learn.microsoft.com/en-us/azure/azure-monitor/app/javascript?tabs=snippet#configuration
            config: {
                connectionString: props.connectionString,
                maxBatchInterval: 0,
                disableFetchTracking: false,  //Fetch requests aren't autocollected.
                disableAjaxTracking: true,    //Ajax calls aren't autocollected.
                extensions: [AppInsightsLogListener.reactPlugin ],
                extensionConfig: {
                    [AppInsightsLogListener.reactPlugin.identifier]: { history: browserHistory }
                }
            }
        });

        appInsights.loadAppInsights();
        appInsights.context.application.ver = props.version;
        return appInsights;
    }

    public static get ReactPlugin(): ReactPlugin {
        if (!AppInsightsLogListener.reactPlugin ) {
            AppInsightsLogListener.reactPlugin  = new ReactPlugin();
        }
        return AppInsightsLogListener.reactPlugin ;
    }
    //https://learn.microsoft.com/en-us/azure/azure-monitor/app/api-custom-events-metrics#trackevent
    public trackEvent(name: string): void {
        if (AppInsightsLogListener.appInsights )
            AppInsightsLogListener.appInsights.trackEvent({ name: name });
    }

    public log(entry: ILogEntry): void {

        const parseMessage = (entry: ILogEntry):string=>{
            const msg: string[] = [];

            msg.push(entry.message);

            if (entry.data) {
                try {
                    msg.push('Data: ' + JSON.stringify(entry.data));
                } catch (e) {
                    msg.push(`Data: Error in stringify of supplied data ${e}`);
                }
            }
            return msg.join(' | ');
        }

        const msg = parseMessage(entry);

        if (entry.level === LogLevel.Off) {
            // No log required since the level is Off
            return;
        }

        if (AppInsightsLogListener.appInsights )
            switch (entry.level) {
                case LogLevel.Verbose:
                    AppInsightsLogListener.appInsights.trackTrace({ message: msg, severityLevel: SeverityLevel.Verbose }, {customProps:"CustomProps"});
                    break;
                case LogLevel.Info:
                    AppInsightsLogListener.appInsights.trackTrace({ message: msg, severityLevel: SeverityLevel.Information }, { customProps: "CustomProps" });
                    console.log({ customProps: "CustomProps", Message: msg });
                    break;
                case LogLevel.Warning:
                    AppInsightsLogListener.appInsights.trackTrace({ message: msg, severityLevel: SeverityLevel.Warning }, { customProps: "CustomProps" });
                    console.warn({ customProps: "CustomProps", Message: msg });
                    break;
                case LogLevel.Error:
                    AppInsightsLogListener.appInsights.trackException({ error: new Error(msg), severityLevel: SeverityLevel.Error });
                    console.error({ customProps:"CustomProps", Message: msg });
                    break;
            }
    }
}