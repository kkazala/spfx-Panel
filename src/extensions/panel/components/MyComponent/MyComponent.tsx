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

        const selectedRow = head(props.selectedRows);

        setIsSubmitted(selectedRow.getValueByName("Submitted") == "Yes" ? true : false);
        setItemID(selectedRow.getValueByName("ID"));
        setListGuid(props.context.listView.list.guid);
        
    }, []);

    
    async function updateListItem(): Promise<boolean> {

        try {
            const sp = spfi().using(SPFx(props.context));
            const items: any[] = await sp.web.lists.getById(listGuid).items.top(1).filter(`ID eq '${itemId}'`)();

            if (items.length > 0) {
                const updatedItem = await sp.web.lists.getById(listGuid).items.getById(items[0].Id).update({
                    Submitted: isSubmitted,
                    // Submitted: "yes" // USE IT TO THROW ERROR,
                });
                return true;
            }
            else
                return false;

        } catch (error) {
            Logger.error(error);
        }
    }
    
    function onToggleChange(_ev: React.MouseEvent<HTMLElement>, checked?: boolean): void {
        setIsSubmitted(checked);
        setRefreshPage.setTrue();
    }

    async function onFormSubmitted(): Promise<void> {

        setFormDisabled();
        const result: boolean = await updateListItem();
        // setFormEnabled();

        console.log(result);
        if (result) {
            setStatusTxt("Page will refresh automatically after you close this panel.");
            setStatusType(MessageBarType.success);
        }
    }

    function onPanelDismissed(): void {
        if (refreshPage && props.panelConfig.onDismiss != undefined) {
            props.panelConfig.onDismiss();
        }
    }
    
    return <StatefulPanel
        title={props.panelConfig.title}
        panelTop={props.panelConfig.panelTop}
        shouldOpen={props.panelConfig.shouldOpen}
        onDismiss={onPanelDismissed}
    >
        {statusTxt && refreshPage &&
            <MessageBar messageBarType={statusType} isMultiline={true} dismissButtonAriaLabel="x" onDismiss={() => setStatusTxt(null) }>{statusTxt}</MessageBar>
        }
        <Toggle
            label="Trip report submitted:"
            inlineLabel
            onChange={onToggleChange}
            defaultChecked={isSubmitted}
            onText="Yes"
            offText="No"
            disabled={formDisabled}
        />
        <PrimaryButton text="OK" onClick={onFormSubmitted} allowDisabledFocus disabled={formDisabled}  />

    </StatefulPanel>;
}