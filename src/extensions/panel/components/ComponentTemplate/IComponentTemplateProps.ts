import { RowAccessor } from "@microsoft/sp-listview-extensibility";
import { IStatefulPanelProps } from "../StatefulPanel/IStatefulPanelProps";

export interface IComponentTemplateProps { 
    selectedRows: readonly RowAccessor[];
    context: any;
    panelConfig: IStatefulPanelProps;
    onChange?: () => void;  
}