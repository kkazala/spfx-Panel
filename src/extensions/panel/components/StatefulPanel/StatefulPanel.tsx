import { IPanelStyles, MessageBar, MessageBarType, Panel, PanelType } from "@fluentui/react";
import { useBoolean } from '@fluentui/react-hooks';
import { loadStyles } from '@microsoft/load-themed-styles';
import * as React from "react";
import { IStatefulPanelProps } from "./IStatefulPanelProps";
import { Logger, LogLevel } from "@pnp/logging";
import { reactPlugin } from "../../../utils/AppInsights";
import AppInsightsErrorBoundary from "../AppInsightsErrorBoundary/AppInsightsErrorBoundary";

export default function StatefulPanel(props: React.PropsWithChildren<IStatefulPanelProps>): JSX.Element {
    const IframePanelStyles: Partial<IPanelStyles> = { root: { top: props.panelTop } };
    const [isOpen, { setTrue: setPanelOpen, setFalse: setPanelClosed }] = useBoolean(false);

    React.useEffect(() => {
        loadStyles('panel');
    }, []);

    React.useEffect(() => {
        if (props.shouldOpen && !isOpen ) {
            setPanelOpen();
        }
    }, [props.shouldOpen]);

    const _onPanelClosed = ():void => {
        props.shouldOpen = false;
        setPanelClosed();

        if (props.onDismiss !== undefined) {
            props.onDismiss();
        }
    };
    const _errorFallback = (error:Error, info: { componentStack: string }): JSX.Element => {
        Logger.error(error);

        return <MessageBar messageBarType={MessageBarType.error} isMultiline={true} dismissButtonAriaLabel="Close" >{error}</MessageBar>;
    };

    return (
        <Panel
        className='od-Panel'
        headerText={props.title}
        isOpen={isOpen}
        type={PanelType.medium}
        isLightDismiss={false}
        styles={IframePanelStyles}
        // key={ props.uniqueKey}
        onDismiss={_onPanelClosed}>
        {/* Ensure there are children to render, otherwise ErrorBoundary throws error */}
        {props.children &&
        <AppInsightsErrorBoundary onError={_errorFallback} appInsights={reactPlugin}>
            <>
                {props.children}
            </>
        </AppInsightsErrorBoundary>
        
        }
        </Panel>
    );
}
