import * as React from "react";
import { TextField } from '@fluentui/react/lib/TextField';

export interface ListItem {
	id: number;
	name: string;
	placeholder: string;
	icon?: string;
}

export interface ListProps {
	message: string;
}

export default class List extends React.Component<ListProps> {
	render() {
		const { message, children } = this.props;

		return (
			<main className="ms-welcome__main">
				<div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 5, padding: "1rem 0" }}>
					{children}
				</div>
			</main>
		)
	}
}
