import { PrimaryButton, CompoundButton } from '@fluentui/react/lib/Button';
import * as React from "react";

export default function RequestedButtons() {
	const [loading, setLoading] = React.useState(true);
	const [data, setData] = React.useState<string[]>();

	React.useEffect(() => {
		fetch("https://wserv-qa.proeducative.com/program/getProducDetailById", {
			method: "POST",
			headers: {
				"Content-Type": "application/x-www-form-urlencoded",
				"Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpYXQiOjE2NzM0NzUwNDgsImV4cCI6NjMyMTcwMDcyMDg2LCJrZXkiOiJBSDNpSDE2UFhTMiJ9.F9OWhE0FcV96AqI0SYIkZz3JLR81ReRjJG7Z1zdG2w4"
			},
			body: new URLSearchParams([["id", "8240"]])
		})
			.then(raw => raw.json()).then(res => {
				setData(Object.keys(res.data.docentes[0])); //docentes[0] // features[0] // product
				setLoading(false);
			})
	}, []);

	if (loading) return <span>Cargando...</span>


	// params
	// master = master=""
	// client =
 
	return (
		<div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "5px" }}>
			{data.map(el => (
				<CompoundButton
					primary
					secondaryText={"[image]"}
					text={`{{${el}}}`}
					onClick={() => insertVariable(el)}
				/>
			))}
		</div>
	)
}
// interface variable {[{
// 	name: string,
// 	type: string
// }]
async function insertVariable (variable: string) {
	Office.context.document.setSelectedDataAsync(`{{${variable}}}`);
}
