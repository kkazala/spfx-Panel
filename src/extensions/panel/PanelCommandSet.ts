import { override } from '@microsoft/decorators';
import { loadStyles } from '@microsoft/load-themed-styles';
import {
  BaseListViewCommandSet, Command, IListViewCommandSetExecuteEventParameters, ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import {
  ConsoleListener, Logger
} from "@pnp/logging";
import { sp } from "@pnp/sp";
import * as React from "react";
import * as ReactDOM from 'react-dom';
import { IMyComponentProps } from './components/MyComponent/IMyComponentProps';
import MyComponent from './components/MyComponent/MyComponent';
import { IStatefulPanelProps } from './components/StatefulPanel/IStatefulPanelProps';
import StatefulPanel from './components/StatefulPanel/StatefulPanel';

export interface IPanelCommandSetProperties {
  sampleTextOne: string;
  sampleTextTwo: string;
  logLevel?: number;
}

const LOG_SOURCE: string = 'PanelCommandSet';

export default class PanelCommandSet extends BaseListViewCommandSet<IPanelCommandSetProperties> {
  private panelPlaceHolder: HTMLDivElement = null;
  private panelTop: number;

  private _showComponent = (props: IMyComponentProps): void => { 
    ReactDOM.render(React.createElement(MyComponent, props), this.panelPlaceHolder);
  }

  private _showPanel = (props: IStatefulPanelProps): void => { 
    ReactDOM.render(React.createElement(StatefulPanel, props), this.panelPlaceHolder);
  }

  private _setLogger = () => {
    Logger.subscribe(new ConsoleListener());
    if (this.properties.logLevel && this.properties.logLevel in [0, 1, 2, 3, 99]) {
      Logger.activeLogLevel = this.properties.logLevel;
    }
    Logger.write(`${LOG_SOURCE} Initialized PanelCommandSet`);  
    Logger.write(`${LOG_SOURCE} Activated Initialized with properties:`);  
    Logger.write(`${LOG_SOURCE} ${JSON.stringify(this.properties, undefined, 2)}`);
  }

  @override
  public onInit(): Promise<void> {
    sp.setup(this.context);
    loadStyles('panel');
    this._setLogger();
    this.panelTop = document.querySelector("#SuiteNavWrapper").clientHeight;
    this.panelPlaceHolder = document.body.appendChild(document.createElement("div"));

    this.context.listView.listViewStateChangedEvent.add(this, this.onListViewUpdatedv2);
    return Promise.resolve();
  }

  public onListViewUpdatedv2(args: ListViewStateChangedEventArgs): void{
    const compareTwoCommand: Command = this.tryGetCommand('COMMAND_2');
    if (compareTwoCommand) { 
      console.log(this.context.listView.selectedRows.length);
      
      compareTwoCommand.disabled = this.context.listView.selectedRows.length == 0;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    
    switch (event.itemId) {
      case 'COMMAND_1':
      this.raiseOnChange();
        this._showPanel({
          shouldOpen:true,
          title: this.properties.sampleTextOne,
          panelTop:this.panelTop
        });
        break;
      case 'COMMAND_2':
        this._showComponent({
          panelConfig: {
            panelTop:this.panelTop,
            shouldOpen:true,
            title: this.properties.sampleTextTwo,
          },
          selectedRows: event.selectedRows,
          context: this.context,
          onRefresh: this.raiseOnChange()
        });
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}

