import { MessageBar, MessageBarType, PrimaryButton, Toggle } from '@fluentui/react';
import { useBoolean } from '@fluentui/react-hooks';
import {
    FunctionListener,
    ILogEntry,
    Logger,
    LogLevel
} from "@pnp/logging";
import "@pnp/sp/items";
import { IItems } from '@pnp/sp/items';
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import { head } from 'lodash';
import * as React from "react";
import StatefulPanel from "../StatefulPanel/StatefulPanel";
import { IMyComponentProps } from "./IMyComponentProps";

export default function MyComponent(props: IMyComponentProps): JSX.Element {
    const [refreshPage, setRefreshPage] = useBoolean(false);
    const [formDisabled, setFormDisabled] = useBoolean(false);
    const [isSubmitted, setIsSubmitted] = React.useState<boolean>(null);
    const [itemId, setItemID] = React.useState(null);
    const [listName, setListName] = React.useState(null);

    const [statusTxt, setStatusTxt] = React.useState<string>(null);
    const [statusType, setStatusType] = React.useState<MessageBarType>(null);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
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

        setIsSubmitted(selectedRow.getValueByName("Submitted") === "Yes" ? true : false);
        setItemID(selectedRow.getValueByName("ID"));
        setListName(props.listName);

    }, []);


    async function updateListItem(): Promise<boolean> {

        try {
            const items: IItems = await props.spfiContext.web.lists.getByTitle(listName).items.top(1).filter(`ID eq '${itemId}'`)();

            if (items.length > 0) {
                await props.spfiContext.web.lists.getByTitle(listName).items.getById(items[0].Id).update({
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

        setFormDisabled.setTrue();
        const result: boolean = await updateListItem();
        // setFormEnabled();

        console.log(result);
        if (result) {
            setStatusTxt("Page will refresh automatically after you close this panel.");
            setStatusType(MessageBarType.success);
        }
    }

    function onPanelDismissed(): void {
        if (refreshPage && props.panelConfig.onDismiss !== undefined) {
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