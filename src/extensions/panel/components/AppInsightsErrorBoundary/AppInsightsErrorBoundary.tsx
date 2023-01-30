/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import { SeverityLevel } from "@microsoft/applicationinsights-common";
import { ReactPlugin } from "@microsoft/applicationinsights-react-js";

export interface IAppInsightsErrorBoundaryProps {
    appInsights: ReactPlugin
    onError: React.ComponentType<any>
    children: React.ReactElement
}

export interface IAppInsightsErrorBoundaryState {
    hasError: boolean
}

export default class AppInsightsErrorBoundary extends React.Component<IAppInsightsErrorBoundaryProps, IAppInsightsErrorBoundaryState> {
    constructor(props: IAppInsightsErrorBoundaryProps | Readonly<IAppInsightsErrorBoundaryProps>) {
        super(props);
        this.state = { hasError: false };
    }

    componentDidCatch(error: Error, errorInfo: React.ErrorInfo):void {
        this.setState({ hasError: true });
        this.props.appInsights.trackException({
            error: error,
            exception: error,
            severityLevel: SeverityLevel.Error,
            properties: errorInfo
        });
    }

    render(): React.ReactElement<any, string | React.JSXElementConstructor<any>> & React.ReactNode {
        if (this.state.hasError) {
            const { onError } = this.props;
            return React.createElement(onError);
        }

        return this.props.children;
    }
}