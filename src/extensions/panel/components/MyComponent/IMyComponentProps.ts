import { RowAccessor } from "@microsoft/sp-listview-extensibility";
import { SPFI } from "@pnp/sp";
import { IStatefulPanelProps } from "../StatefulPanel/IStatefulPanelProps";

export interface IMyComponentProps {
    selectedRows: readonly RowAccessor[];
    spfiContext: SPFI;
    listName: string;
    panelConfig: IStatefulPanelProps;
}