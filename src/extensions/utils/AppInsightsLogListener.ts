
import { AnalyticsPlugin } from '@microsoft/applicationinsights-analytics-js';
import { SeverityLevel } from '@microsoft/applicationinsights-web';
import { ILogEntry, ILogListener, LogLevel } from "@pnp/logging";


export class AppInsightsLogListener implements ILogListener {
    private appInsights: AnalyticsPlugin;
    constructor(appInsights: AnalyticsPlugin) {
        this.appInsights = appInsights;
    }

    public log(entry: ILogEntry): void {

        const parseMessage = (entry: ILogEntry): string => {
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


        if (entry.level === LogLevel.Off) {
            // No log required since the level is Off
            return;
        }
        const msg = parseMessage(entry);

        if (this.appInsights){
            switch (entry.level) {
                case LogLevel.Verbose:
                    this.appInsights.trackTrace({ message: msg, severityLevel: SeverityLevel.Verbose }, { customProps: "CustomProps" });
                    break;
                case LogLevel.Info:
                    this.appInsights.trackTrace({ message: msg, severityLevel: SeverityLevel.Information }, { customProps: "CustomProps" });
                    break;
                case LogLevel.Warning:
                    this.appInsights.trackTrace({ message: msg, severityLevel: SeverityLevel.Warning }, { customProps: "CustomProps" });
                    break;
                case LogLevel.Error:
                    this.appInsights.trackException({ error: new Error(msg), severityLevel: SeverityLevel.Error });
                    break;
            }
        }
    }
}

