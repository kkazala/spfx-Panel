import { override } from '@microsoft/decorators';
import {
  BaseListViewCommandSet, Command, IListViewCommandSetExecuteEventParameters, ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { ConsoleListener, Logger } from '@pnp/logging';
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

  @override
  public onInit(): Promise<void> {
    
    const _setLogger = (logSource: string, properties: any, logLevel?: number) => {
      
      Logger.subscribe(new (ConsoleListener as any)(logSource));
      if (logLevel && logLevel in [0, 1, 2, 3, 99]) {
        Logger.activeLogLevel = logLevel;
      }
      Logger.write(`Initialized PanelCommandSet`);
      Logger.write(`Activated Initialized with properties:`);
      Logger.writeJSON(properties);
    }
  
    const _setCommandsHidden = () => {
      const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
      if (compareOneCommand) {
        compareOneCommand.visible = false;
      }
      const compareTwoCommand: Command = this.tryGetCommand('COMMAND_2');
      if (compareTwoCommand) {
        compareTwoCommand.visible = false;
      }
    }
  
    _setLogger(LOG_SOURCE, this.properties, this.properties.logLevel);

    this.panelTop = document.querySelector("#SuiteNavWrapper").clientHeight;
    this.panelPlaceHolder = document.body.appendChild(document.createElement("div"));

    this.context.listView.listViewStateChangedEvent.add(this, this.onListViewUpdatedv2);
    
  // set all commands: visible = false, because
  // there's an issue detecting list context during onInit, see https://github.com/SharePoint/sp-dev-docs/issues/7795 
    _setCommandsHidden();

    return Promise.resolve();
  }

  // Triggered when row(s) un/selected
  public onListViewUpdatedv2(args: ListViewStateChangedEventArgs): void{

    Logger.write("onListViewUpdatedv2");

    const isCorrectList = (this.context.listView.list.title == "Travel requests") ? true : false;
    const itemSelected = this.context.listView.selectedRows && this.context.listView.selectedRows.length == 1;
    
    let raiseOnChange: boolean = false;
    
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand && (compareOneCommand.visible != isCorrectList )) {
      compareOneCommand.visible = isCorrectList;
      raiseOnChange = true;
    }
    
    const compareTwoCommand: Command = this.tryGetCommand('COMMAND_2');
    if (compareTwoCommand && (compareTwoCommand.visible!= (isCorrectList && itemSelected))) {
      compareTwoCommand.visible = isCorrectList && itemSelected;
      raiseOnChange = true;
    }

    // NOTE: use it carefully, frequent calls can lead to low performance of the list
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
      
      const ChangeToken = Date.now(); //timestamp in milliseconds
      ReactDOM.render(React.createElement(MyComponent, { ...props, key: ChangeToken }), this.panelPlaceHolder);

      //about react components with key: 
      //  https://dev.to/francodalessio/understanding-the-importance-of-the-key-prop-in-react-3ag7
      //  https://kentcdodds.com/blog/understanding-reacts-key-prop
      
    }

    const _showPanel = (props: IStatefulPanelProps): void => { 
      ReactDOM.render(React.createElement(StatefulPanel, props), this.panelPlaceHolder);
    }

    const _refreshList = (): void => { 
      Logger.write(`Refreshing list view`);  
      location.reload();
    }
  
    switch (event.itemId) {
      case 'COMMAND_1':
        _showPanel({
          shouldOpen:true,
          title: this.properties.sampleTextOne,
          panelTop:this.panelTop
        });
        break;
      case 'COMMAND_2':
        _showComponent({
          panelConfig: {
            panelTop:this.panelTop,
            shouldOpen:true,
            title: this.properties.sampleTextTwo,
            onDismiss: _refreshList
          },
          context: this.context,
          selectedRows: event.selectedRows
        });
        break;
      default:
        throw new Error('Unknown command');
    }
  }
}

