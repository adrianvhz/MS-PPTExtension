import { PrimaryButton, CompoundButton } from '@fluentui/react/lib/Button';
import * as React from "react";
// import generateToken from '../../lib/generateToken';

export default function RequestedButtons() {
	const [loading, setLoading] = React.useState(true);
	const [data, setData] = React.useState<any[]>();

	React.useEffect(() => {
		const xhr = new XMLHttpRequest();
		xhr.open("POST", "https://idcloudsystem.com/api/products/variables");
		// xhr.setRequestHeader("Authorization", "Bearer " + window.localStorage.getItem("tk") || generateToken());
		xhr.setRequestHeader("Authorization", "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpYXQiOjE2NzQwNTk3MTYsImV4cCI6MTczNzMwNDc0Njc2NCwia2V5IjoiaFZtWXEzdDYifQ.qRVUF9wcJp7YNiOk-Saknf6e9trRuhD3uScm6y2VLkk");
		xhr.onload = () => {
			const res = JSON.parse(xhr.response);
			setData(res.data.variables);
			setLoading(false);
		}
		xhr.send();

		// fetch("https://idcloudsystem.com/api/products/30/detail", {
		// 	method: "POST",
		// 	headers: {
		// 		"Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpYXQiOjE2NzQwNTk3MTYsImV4cCI6MTY3NDE0NjExNiwia2V5IjoidXNlcnBpZWNlIn0.AX6MoJVyeG54noHdy3GglwsXf-aY7J8TijxfPosIeWM"
		// 	},
		// })
		// 	.then(raw => raw.json()).then(res => {
		// 		console.log(res)
		// 		// setData(Object.keys(res.data.docentes[0])); //docentes[0] // features[0] // product
		// 		// setLoading(false);
		// 	})
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

	// Office.context.document.getFilePropertiesAsync(null, (x) => {
	// 	console.log(x.value);
	// });
}
