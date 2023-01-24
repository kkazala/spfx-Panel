import { ReactPlugin } from "@microsoft/applicationinsights-react-js";


export interface IStatefulPanelProps {
    title: string;
    shouldOpen: boolean;
    panelTop: number;
    onDismiss?: () => void;
    uniqueKey?: string;
    reactPlugin: ReactPlugin;
}