import {
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Toggle,
} from "@fluentui/react";
import { useBoolean } from "@fluentui/react-hooks";
import {
  AppInsightsContext,
  AppInsightsErrorBoundary,
} from "@microsoft/applicationinsights-react-js";
import { FunctionListener, ILogEntry, Logger, LogLevel } from "@pnp/logging";
import "@pnp/sp/items";
import { IItems } from "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import * as strings from "PanelCommandSetStrings";
import * as React from "react";
import { reactPlugin } from "../../../utils/AppInsights";
import { handleError } from "../../../utils/ErrorHandler";
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
    if (props.selectedRows.length > 0) {
      const selectedRow = props.selectedRows[0];
      setIsSubmitted(
        selectedRow.getValueByName("Submitted") === "Yes" ? true : false
      );
      setItemID(selectedRow.getValueByName("ID"));
    }
    setListName(props.listName);
  }, []);

  const _errorFallback = (
    error: Error,
    info: { componentStack: string }
  ): JSX.Element => {
    Logger.error(error);
    Logger.write(info.componentStack, LogLevel.Error);
    return (
      <MessageBar
        messageBarType={MessageBarType.error}
        isMultiline={true}
        dismissButtonAriaLabel="Close"
      >
        {error}
      </MessageBar>
    );
  };
  const updateListItem = async (): Promise<boolean> => {
    let success: boolean = false;
    try {
      const items: IItems = await props.spfiContext.web.lists
        .getByTitle(listName)
        .items.top(1)
        .filter(`ID eq '${itemId}'`)();

      if (items.length > 0) {
        await props.spfiContext.web.lists
          .getByTitle(listName)
          .items.getById(items[0].Id)
          .update({
            Submitted: isSubmitted,
            // Submitted: "yes" // USE IT TO THROW ERROR,
          });

        success = true;
      }
      return success;
    } catch (error) {
      // Parse error message in case it's a HttpRequestError error
      // log using Logger
      await handleError(error);
      return success;
    }
  };

  const onToggleChange = (
    _ev: React.MouseEvent<HTMLElement>,
    checked?: boolean
  ): void => {
    setIsSubmitted(checked);
    setRefreshPage.setTrue();
  };

  const onFormSubmitted = async (): Promise<void> => {
    setFormDisabled.setTrue();
    const result: boolean = await updateListItem();
    if (props.onCompleted !== undefined) {
      props.onCompleted(result);
    }

    //if form values changed and item succesfully updated
    if (result && refreshPage) {
      setStatusTxt(strings.lblPageWillRefresh);
      setStatusType(MessageBarType.success);
    }
  };

  const onPanelDismissed = (): void => {
    if (refreshPage && props.panelConfig.onDismiss !== undefined) {
      props.panelConfig.onDismiss();
    }
  };

  return (
    <AppInsightsContext.Provider value={reactPlugin}>
      <AppInsightsErrorBoundary
        onError={_errorFallback}
        appInsights={reactPlugin}
      >
        <StatefulPanel
          title={props.panelConfig.title}
          panelTop={props.panelConfig.panelTop}
          shouldOpen={props.panelConfig.shouldOpen}
          onDismiss={onPanelDismissed}
        >
          {statusTxt && (
            <MessageBar
              messageBarType={statusType}
              isMultiline={true}
              dismissButtonAriaLabel="x"
              onDismiss={() => setStatusTxt(null)}
            >
              {statusTxt}
            </MessageBar>
          )}
          <Toggle
            label={strings.lblConfirm}
            inlineLabel
            onChange={onToggleChange}
            defaultChecked={isSubmitted}
            onText={strings.lblYes}
            offText={strings.lblNo}
            disabled={formDisabled}
          />
          <PrimaryButton
            text={strings.btnSubmit}
            onClick={onFormSubmitted}
            allowDisabledFocus
            disabled={formDisabled}
          />
        </StatefulPanel>
      </AppInsightsErrorBoundary>
    </AppInsightsContext.Provider>
  );
}
