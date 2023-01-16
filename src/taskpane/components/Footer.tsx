import * as React from "react";

export interface FooterProps {
	title: string;
	logo: string;
	message: string;
}

export default function Footer(props: FooterProps) {
	const { title, logo, message } = props
	
	return (
		<footer>
			<p className="ms-fontWeight-light" style={{ textAlign: "center" }}>© PD 2023 - { "{ code }" }</p>{/* b691bf */}
		</footer>
	)
}
