import { MessageBar, MessageBarType, PrimaryButton } from '@fluentui/react';
import { useBoolean } from '@fluentui/react-hooks';
import {
    FunctionListener,
    ILogEntry,
    Logger,
    LogLevel
} from "@pnp/logging";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import * as React from "react";
import StatefulPanel from "../StatefulPanel/StatefulPanel";
import { IComponentTemplateProps } from './IComponentTemplateProps';


export default function ComponentTemplate(props: IComponentTemplateProps) {
    const [refreshPage, setRefreshPage] = useBoolean(false);
    const [formDisabled, { setTrue: setFormDisabled, setFalse:setFormEnabled }] = useBoolean(false);
    const [isSubmitted, setIsSubmitted] = React.useState<boolean>(null);
    const [itemId, setItemID] = React.useState(null);
    const [listGuid, setListGuid] = React.useState(null);

    const [statusTxt, setStatusTxt] = React.useState<string>(null);
    const [statusType, setStatusType] = React.useState<MessageBarType>(null);

    const funcListener = new (FunctionListener as any)((entry: ILogEntry) => {
        switch (entry.level) {
            case LogLevel.Error:
                setStatusTxt(entry.message);
                setStatusType(MessageBarType.error);
                break;
            case LogLevel.Warning:
                setStatusTxt(entry.message);
                setStatusType(MessageBarType.warning);
                break;
        }
    });

    React.useEffect(() => {
        Logger.subscribe(funcListener);
    }, []);

    const someFunction = async() => {

        try {
            const sp = spfi().using(SPFx(props.context));

            return true;

        } catch (error) {
            Logger.error(error);
        }
    }

    const _onFormSubmitted = async() => {

        setFormDisabled();
        const result = await someFunction();
        setFormEnabled();

        if (result && refreshPage) {
            setStatusTxt("OK");
            setStatusType(MessageBarType.success);
            props.onChange();
        }
    };

    function _handleStatusMsgChange(message: string, messageBarType: MessageBarType): void {
        setStatusTxt(message);
        setStatusType(messageBarType);
    }

    return <StatefulPanel
        title={props.panelConfig.title}
        panelTop={props.panelConfig.panelTop}
        shouldOpen={props.panelConfig.shouldOpen}
        onDismiss={props.panelConfig.onDismiss}
        reactPlugin={props.panelConfig.reactPlugin}
    >
        {statusTxt &&
            <MessageBar messageBarType={statusType} isMultiline={true} dismissButtonAriaLabel="x" onDismiss={() => _handleStatusMsgChange(null, null)}>{statusTxt}</MessageBar>
        }

        <PrimaryButton text="OK" onClick={_onFormSubmitted} allowDisabledFocus disabled={formDisabled}  />

    </StatefulPanel>;
}