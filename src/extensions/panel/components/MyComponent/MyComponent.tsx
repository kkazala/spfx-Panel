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
import * as strings from 'PanelCommandSetStrings';
import * as React from "react";
import { handleError } from '../../../utils/ErrorHandler';
import StatefulPanel from "../StatefulPanel/StatefulPanel";
import { IMyComponentProps } from "./IMyComponentProps";

export default function MyComponent(props: IMyComponentProps): JSX.Element {
    //#region const
    const [refreshPage, setRefreshPage] = useBoolean(false);
    const [formDisabled, setFormDisabled] = useBoolean(false);
    const [isSubmitted, setIsSubmitted] = React.useState<boolean>(null);
    const [itemId, setItemID] = React.useState(null);
    const [listName, setListName] = React.useState(null);

    const [statusTxt, setStatusTxt] = React.useState<string>(null);
    const [statusType, setStatusType] = React.useState<MessageBarType>(null);
    //#endregion

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

        let success:boolean=false;
        try {
            const items: IItems = await props.spfiContext.web.lists.getByTitle(listName).items.top(1).filter(`ID eq '${itemId}'`)();

            if (items.length > 0) {
                await props.spfiContext.web.lists.getByTitle(listName).items.getById(items[0].Id).update({
                    Submitted: isSubmitted,
                    // Submitted: "yes" // USE IT TO THROW ERROR,
                });

                success= true;
            }
            return success;

        } catch (error) {
            handleError(error);
            return success;
        }
    }

    function onToggleChange(_ev: React.MouseEvent<HTMLElement>, checked?: boolean): void {
        setIsSubmitted(checked);
        setRefreshPage.setTrue();
    }

    async function onFormSubmitted(): Promise<void> {

        setFormDisabled.setTrue();
        const result: boolean = await updateListItem();
        if(props.onCompleted!== undefined){
            props.onCompleted(result);
        }

        //if form values changed and item succesfully updated
        if (result && refreshPage ) {
            setStatusTxt(strings.lblPageWillRefresh);
            setStatusType(MessageBarType.success);
        }
    }

    function onPanelDismissed(): void {
        if (refreshPage && props.panelConfig.onDismiss !== undefined) {
            props.panelConfig.onDismiss();
        }
    }

    return (
        <StatefulPanel
            title={props.panelConfig.title}
            panelTop={props.panelConfig.panelTop}
            shouldOpen={props.panelConfig.shouldOpen}
            onDismiss={onPanelDismissed}
        >
            {statusTxt && 
                <MessageBar messageBarType={statusType} isMultiline={true} dismissButtonAriaLabel="x" onDismiss={() => setStatusTxt(null) }>{statusTxt}</MessageBar>
            }
            <Toggle
                label={strings.lblConfirm}
                inlineLabel
                onChange={onToggleChange}
                defaultChecked={isSubmitted}
                onText={strings.lblYes}
                offText={strings.lblNo}
                disabled={formDisabled}
            />
            <PrimaryButton text={strings.btnSubmit} onClick={onFormSubmitted} allowDisabledFocus disabled={formDisabled}  />
        </StatefulPanel>);
}
