import { ReactPlugin, withAITracking } from "@microsoft/applicationinsights-react-js";
import { ApplicationInsights } from "@microsoft/applicationinsights-web";
import { createBrowserHistory } from "history";
import { ComponentType, memo } from "react";

export interface IAppInsightsProps {
    connectionString: string;
    version: string; //this.manifest.version
}
const browserHistory = createBrowserHistory({});
export const reactPlugin = new ReactPlugin();
export class AppInsights extends ApplicationInsights {
    constructor(props: IAppInsightsProps){
        super({config:{
            connectionString: props.connectionString,
            maxBatchInterval: 0,
            disableFetchTracking: false,  //Fetch requests aren't autocollected.
            disableAjaxTracking: true,    //Ajax calls aren't autocollected.
            extensions: [reactPlugin],
            extensionConfig: {
                [reactPlugin.identifier]: { history: browserHistory }
            }
        }})
        this.context.application.ver= props.version;
      // this.setAuthenticatedUserContext

    }
}

export function wrap<T extends ComponentType<unknown>>(
    component: T,
    componentName?: string,
    className?: string
): T {
    return withAITracking(reactPlugin, memo(component), componentName, className) as T
}

