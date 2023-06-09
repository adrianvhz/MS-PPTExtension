import * as React from "react";

export interface NotificationBarProps {
	notification: {
		message: string;
		error: boolean;
	}
}

export default class NotificationBar extends React.Component<NotificationBarProps> {
	render() {
		const {
			notification: {
				message, error
			}
		} = this.props;

		return (
			<div>
				<h4 style={{ marginBottom: ".6em" }}><u>Notifications</u></h4>
				<p style={{ color: error ? "#cd5c5c" : "initial", margin: 0 }}>{message}</p>
			</div>
		)
	}
}
