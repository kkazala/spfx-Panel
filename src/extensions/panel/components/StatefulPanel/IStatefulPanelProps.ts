export interface IStatefulPanelProps {
    title: string;
    shouldOpen: boolean;
    panelTop: number;
    onDismiss?: () => void;
    uniqueKey?: string;
}