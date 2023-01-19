import { PrimaryButton, CompoundButton } from '@fluentui/react/lib/Button';
import * as React from "react";
import generateToken from '../../lib/generateToken';

export default function RequestedButtons() {
	const [loading, setLoading] = React.useState(true);
	const [data, setData] = React.useState<any[]>();

	React.useEffect(() => {
		const xhr = new XMLHttpRequest();
		xhr.open("POST", "https://idcloudsystem.com/api/products/variables");
		xhr.setRequestHeader("Authorization", "Bearer " + window.localStorage.getItem("tk") || generateToken());
		xhr.onload = () => {
			const res = JSON.parse(xhr.response);
			setData(res.data.variables);
			setLoading(false);
		}
		xhr.send();
	}, []);

	if (loading) return <span>Cargando...</span>

	console.log(data)

	return (
		<div
			style={{
				display: "grid",
				// gridTemplateColumns: "1fr 1fr",
				gridTemplateColumns: "1fr",
				gap: "5px"
			}}>
			{data.map(el => (
				<CompoundButton
					primary
					secondaryText={el.type}
					text={el.name}
					onClick={() => insertVariable(el.name)}
				/>
			))}
		</div>
	)
}

async function insertVariable (variable: string) {
	Office.context.document.setSelectedDataAsync(variable , {
		coercionType: Office.CoercionType.Text
	});
}
