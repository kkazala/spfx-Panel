import { IPanelStyles, Panel, PanelType } from "@fluentui/react";
import { useBoolean } from "@fluentui/react-hooks";
import { loadStyles } from "@microsoft/load-themed-styles";
import * as React from "react";
import { IStatefulPanelProps } from "./IStatefulPanelProps";

export default function StatefulPanel(props: React.PropsWithChildren<IStatefulPanelProps>): JSX.Element {
	const IframePanelStyles: Partial<IPanelStyles> = {
		root: { top: props.panelTop },
	};
	const [isOpen, { setTrue: setPanelOpen, setFalse: setPanelClosed }] = useBoolean(false);

	React.useEffect(() => {
		loadStyles("panel");
		setPanelOpen();
	}, []);

	const _onPanelClosed = (): void => {
		setPanelClosed();

		if (props.onDismiss !== undefined) {
			props.onDismiss();
		}
	};

	return (
		<Panel className="od-Panel" headerText={props.title} isOpen={isOpen} type={PanelType.medium} isLightDismiss={false} styles={IframePanelStyles} onDismiss={_onPanelClosed}>
			{/* Ensure there are children to render, otherwise ErrorBoundary throws error */}
			{props.children && <>{props.children}</>}
		</Panel>
	);
}
