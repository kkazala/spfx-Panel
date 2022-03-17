import { MessageBar, MessageBarType, PrimaryButton, Toggle } from '@fluentui/react';
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
import { head } from 'lodash';
import * as React from "react";
import StatefulPanel from "../StatefulPanel/StatefulPanel";
import { IMyComponentProps } from "./IMyComponentProps";

export default function MyComponent(props: IMyComponentProps) {
    const [refreshPage, setRefreshPage] = useBoolean(false);
    const [formDisabled, { setTrue: setFormDisabled, setFalse:setFormEnabled }] = useBoolean(false);
    const [isSubmitted, setIsSubmitted] = React.useState<boolean>(null);
    const [itemId, setItemID] = React.useState(null);
    const [listGuid, setListGuid] = React.useState(null);

    const [statusTxt, setStatusTxt] = React.useState<string>(null);
    const [statusType, setStatusType] = React.useState<MessageBarType>(null);


    React.useEffect(() => {
        Logger.subscribe(funcListener);
        Logger.write("MyComponent useEffect[]");

        const selectedRow = head(props.selectedRows);

        setIsSubmitted(selectedRow.getValueByName("Submitted") == "Yes" ? true : false);
        setItemID(selectedRow.getValueByName("ID"));
        setListGuid(props.context.listView.list.guid);
        
    }, []);

    const updateListItem = async() => { 

        const sp = spfi().using(SPFx(props.context));
        const items: any[] = await sp.web.lists.getById(listGuid).items.top(1).filter(`ID eq '${itemId}'`)();

        if (items.length > 0) {
            const updatedItem = await sp.web.lists.getById(listGuid).items.getById(items[0].Id).update({
                Submitted: isSubmitted,
            });
            console.log(updatedItem);
            return true;
        }
        else
            return false;
    }
    
    const _onToggleChange = (ev: React.MouseEvent<HTMLElement>, checked?: boolean) => { 
        setIsSubmitted(checked);
        setRefreshPage.setTrue();
    }

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
    
    const _onFormSubmitted = async() => { 
        Logger.write("MyComponent - form submitted");
        try {
            setFormDisabled();
            const result = await updateListItem();
            setFormEnabled();
            Logger.write(`result && refreshPage: ${result && refreshPage}`);

            if (result && refreshPage) {
                setStatusTxt("OK");
                setStatusType(MessageBarType.success);
                props.onChange();
            }
            
        } catch (error) {
            
        }
  
        

    };
    function handleStatusMsgChange(message: string, messageBarType: MessageBarType): void {
        setStatusTxt(message);
        setStatusType(messageBarType);
    }
    
    return <StatefulPanel
        title={props.panelConfig.title}
        panelTop={props.panelConfig.panelTop}
        shouldOpen={props.panelConfig.shouldOpen}
        onDismiss={props.panelConfig.onDismiss}
    >
        {statusTxt &&
            <MessageBar messageBarType={statusType} isMultiline={true} dismissButtonAriaLabel="x" onDismiss={() => handleStatusMsgChange(null, null)}>{statusTxt}</MessageBar>
        }
        <Toggle
            label="Trip report submitted:"
            inlineLabel
            onChange={_onToggleChange}
            defaultChecked={isSubmitted}
            onText="Yes"
            offText="No"
            disabled={formDisabled}
        />
        <PrimaryButton text="OK" onClick={_onFormSubmitted} allowDisabledFocus disabled={formDisabled}  />

    </StatefulPanel>;
}