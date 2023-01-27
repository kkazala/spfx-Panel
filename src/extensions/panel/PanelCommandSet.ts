import { AnalyticsPlugin } from '@microsoft/applicationinsights-analytics-js';
import { override } from '@microsoft/decorators';
import {
  BaseListViewCommandSet, Command, IListViewCommandSetExecuteEventParameters, ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { ConsoleListener, Logger } from '@pnp/logging';
import { SPFI, spfi, SPFx } from '@pnp/sp';
import * as strings from 'PanelCommandSetStrings';
import * as React from "react";
import * as ReactDOM from 'react-dom';
import { AppInsights } from '../utils/AppInsights';
import { AppInsightsLogListener } from '../utils/AppInsightsLogListener';
import { IMyComponentProps } from './components/MyComponent/IMyComponentProps';
import MyComponent from './components/MyComponent/MyComponent';
import { IStatefulPanelProps } from './components/StatefulPanel/IStatefulPanelProps';
import StatefulPanel from './components/StatefulPanel/StatefulPanel';

export interface IPanelCommandSetProperties {
    logLevel?: number;
    listName:string;
    appInsightsConnectionString?:string;
}
interface IProcessConfigResult{
    visible: boolean;
    disabled:boolean;
    title:string;
}

const LOG_SOURCE: string = 'PanelCommandSet';

export default class PanelCommandSet extends BaseListViewCommandSet<IPanelCommandSetProperties> {
    //#region variables
    private panelPlaceHolder: HTMLDivElement = null;
    private panelTop: number;
    private panelId: string;
    private compId: string;
    private spfiContext: SPFI;
    private appInsights: AnalyticsPlugin;
    //#endregion

  @override
  public onInit(): Promise<void> {

    const _isListRegistered= (this.context.listView.list.title === this.properties.listName) ? true : false;
    
    const _setLogger = (appInsights?:AnalyticsPlugin): void => {

      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      Logger.subscribe(new (ConsoleListener as any)());
      //Application Insights tracking
      if (appInsights !== undefined){
        Logger.subscribe(new AppInsightsLogListener(appInsights));
      }

      if (
        this.properties.logLevel &&
        this.properties.logLevel in [0, 1, 2, 3, 99]
      ) {
        Logger.activeLogLevel = this.properties.logLevel;
      }

      Logger.write(
        `${LOG_SOURCE} Activated Initialized with properties:`
      );
      Logger.write(
        `${LOG_SOURCE} ${JSON.stringify(this.properties, undefined, 2)}`
      );
    };
    const _setPanel = (): void => {
      this.panelTop = document.querySelector("#SuiteNavWrapper").clientHeight;
      this.panelPlaceHolder = document.body.appendChild(document.createElement("div"));
    };
    const _setCommands = ():void => {
      const _setCommandState = (command: Command, config: IProcessConfigResult): void => {
        command.visible = config.visible;
        command.disabled = config.disabled;
        command.title = config.title;
      }

      const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
      if (compareOneCommand) {
        _setCommandState(compareOneCommand, {title: strings.Command1, visible:true, disabled:false});
      }

      const compareTwoCommand: Command = this.tryGetCommand('COMMAND_2');
      if (compareTwoCommand) {
        _setCommandState(compareTwoCommand, { title: strings.Command2, visible: true, disabled: true });
      }

      this.raiseOnChange();
    }

    if (!_isListRegistered){
      return;
    }

    if (this.properties.appInsightsConnectionString){
      this.appInsights= AppInsights(this.properties.appInsightsConnectionString)
      this.appInsights.trackPageView();
      _setLogger(this.appInsights);
    }
    else{
      _setLogger();
    }

    _setPanel();
    _setCommands();
    this.spfiContext = spfi().using(SPFx(this.context)); //https://github.com/pnp/pnpjs/issues/2304
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  // Triggered when row(s) un/selected
  public _onListViewStateChanged(args: ListViewStateChangedEventArgs): void{

    Logger.write("onListViewUpdatedv2");

    const itemSelected = this.context.listView.selectedRows && this.context.listView.selectedRows.length === 1;

    let raiseOnChange: boolean = false;

    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand && compareOneCommand.disabled === itemSelected) {
      compareOneCommand.disabled = !itemSelected;
      raiseOnChange = true;
    }

    const compareTwoCommand: Command = this.tryGetCommand('COMMAND_2');
    if (compareTwoCommand && compareTwoCommand.disabled === itemSelected) {
      compareTwoCommand.disabled = !itemSelected;
      raiseOnChange = true;
    }

    // NOTE: use it with caution; frequent calls can lead to low performance of the list
    // https://github.com/SharePoint/sp-dev-docs/discussions/7375#discussioncomment-2053604
    if (raiseOnChange) {
      this.raiseOnChange();
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {

    const _showComponent = (props: IMyComponentProps): void => {
        //using item's ID as a key will ensure the form is re-rendered when selection is changed
        //BUT if Panel for the Item has been opened (rendered), and
        //   - panel is closed
        //   - sb changed item properties (another browser)
        //   - wait for list to refresh and show the new property
        //   - open Panel: old value becasue panel is not re-rendered
        // const ChangeToken = props.selectedRows[0].getValueByName("ID");

        this.compId = Date.now().toString();
        this.panelPlaceHolder.setAttribute('id', this.compId);
        const element: React.ReactElement<IMyComponentProps> = React.createElement(
            MyComponent, 
            {   
                ...props, 
                key: this.compId 
            }
        );

        ReactDOM.render(element,this.panelPlaceHolder);
    }

    const _showPanel = (props: IStatefulPanelProps): void => {

        this.panelId = Date.now().toString();
        this.panelPlaceHolder.setAttribute('id', this.panelId);

        const element: React.ReactElement<IStatefulPanelProps> = React.createElement(
            StatefulPanel, 
            {   
                ...props, 
                key: this.panelId 
            },
            React.createElement('div', {
                dangerouslySetInnerHTML:{__html: strings.htmlInfo}
            })
        );

        ReactDOM.render(element, this.panelPlaceHolder);
    }

    const _refreshList = (): void => {
        Logger.write(strings.lblRefreshing);
        location.reload();
    }
    const _onCompleted=(success:boolean):void=>{
        if(this.appInsights!==undefined){
            this.appInsights.trackEvent({name: (success)? strings.lblItemUpdate_OK:strings.lblItemUpdate_Err})
        }
    }

    switch (event.itemId) {
      case 'COMMAND_1':
        _showPanel({
          shouldOpen:true,
          title: strings.titleTravelGuidelines,
          panelTop:this.panelTop,
        });
        break;
      case 'COMMAND_2':
        _showComponent({
          panelConfig: {
            panelTop:this.panelTop,
            shouldOpen:true,
            title: strings.titleTravelReport,
            onDismiss: _refreshList,
          },
          spfiContext: this.spfiContext,
          listName: this.context.listView.list.title,
          selectedRows: event.selectedRows,
          onCompleted: _onCompleted
        });
        break;
      default:
        throw new Error('Unknown command');
    }
  }
  
  public onDispose(): void {
    ReactDOM.unmountComponentAtNode(document.getElementById(this.panelId))
    ReactDOM.unmountComponentAtNode(document.getElementById(this.componentId))
  }
}

