import { AnalyticsPlugin } from "@microsoft/applicationinsights-analytics-js";
import { ReactPlugin } from "@microsoft/applicationinsights-react-js";
import { ApplicationInsights } from "@microsoft/applicationinsights-web";
import { createBrowserHistory } from "history";

export interface IAppInsightsProps {
    connectionString: string;
    version: string; //this.manifest.version
}
const browserHistory = createBrowserHistory({});
export const reactPlugin = new ReactPlugin();
export function AppInsights(connString:string): AnalyticsPlugin{
    const  ai = new ApplicationInsights({
        config: {
            connectionString: connString,
            extensions: [reactPlugin],
            extensionConfig: {
            [reactPlugin.identifier]: { history: browserHistory }
            }
        }
    });
    ai.loadAppInsights();
    return ai.appInsights;
}


