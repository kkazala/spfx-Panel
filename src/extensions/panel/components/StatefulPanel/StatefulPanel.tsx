import { IPanelStyles, Panel, PanelType } from "@fluentui/react";
import { useBoolean } from "@fluentui/react-hooks";
import { loadStyles } from "@microsoft/load-themed-styles";
import * as React from "react";
import ReactDOM from "react-dom";
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

	// React.useEffect(() => {
	// 	if (props.shouldOpen && !isOpen) {
	// 		setPanelOpen();
	// 	}
	// }, [props.shouldOpen]);

	const _onPanelClosed = (): void => {
		// props.shouldOpen = false;
		setPanelClosed();
		// unmountComponentAtNode(this.dom)
		ReactDOM.unmountComponentAtNode(document.getElementById(this.key));

		if (props.onDismiss !== undefined) {
			props.onDismiss();
		}
	};

	return (
		<Panel
			className="od-Panel"
			headerText={props.title}
			isOpen={isOpen}
			type={PanelType.medium}
			isLightDismiss={false}
			styles={IframePanelStyles}
			// key={ props.uniqueKey}
			onDismiss={_onPanelClosed}
		>
			{/* Ensure there are children to render, otherwise ErrorBoundary throws error */}
			{props.children && <>{props.children}</>}
		</Panel>
	);
}

// export default function StatefulPanel(props: React.PropsWithChildren<IStatefulPanelProps>): JSX.Element {
// 	const IframePanelStyles: Partial<IPanelStyles> = {
// 		root: { top: props.panelTop },
// 	};
// 	const [isOpen, { setTrue: setPanelOpen, setFalse: setPanelClosed }] = useBoolean(false);

// 	React.useEffect(() => {
// 		loadStyles("panel");
// 	}, []);

// 	React.useEffect(() => {
// 		if (props.shouldOpen && !isOpen) {
// 			setPanelOpen();
// 		}
// 	}, [props.shouldOpen]);

// 	const _onPanelClosed = (): void => {
// 		props.shouldOpen = false;
// 		setPanelClosed();

// 		if (props.onDismiss !== undefined) {
// 			props.onDismiss();
// 		}
// 	};

// 	return (
// 		<Panel
// 			className="od-Panel"
// 			headerText={props.title}
// 			isOpen={isOpen}
// 			type={PanelType.medium}
// 			isLightDismiss={false}
// 			styles={IframePanelStyles}
// 			// key={ props.uniqueKey}
// 			onDismiss={_onPanelClosed}
// 		>
// 			{/* Ensure there are children to render, otherwise ErrorBoundary throws error */}
// 			{props.children && <>{props.children}</>}
// 		</Panel>
// 	);
// }
