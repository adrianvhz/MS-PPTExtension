import * as React from "react";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
// import generateToken from "../../lib/generateToken";

export default function ReplaceVars({ filename }) {
	const [isLoading, setIsLoading] = React.useState(false);

	React.useEffect(() => {
		const alias = filename.split("-")[0];

		setIsLoading(true);
		fetch(`https://idcloudsystem.com/api/products/${alias}/detail`, {
			method: "POST",
			headers: {
				// "Authorization": "Bearer " + window.localStorage.getItem("tk") || generateToken()
				"Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpYXQiOjE2NzQwNTk3MTYsImV4cCI6MTczNzMwNDc0Njc2NCwia2V5IjoiaFZtWXEzdDYifQ.qRVUF9wcJp7YNiOk-Saknf6e9trRuhD3uScm6y2VLkk"
			},
		})
			.then(raw => raw.json()).then(res => {
				replaceAllVars(res.data);
				setIsLoading(false);
			})
	}, []);

	if (isLoading) {
		return (
			<div>
				<Spinner size={SpinnerSize.medium} />
			</div>
		)
	}

	return (
		<p>Done</p>
	)
}

async function replaceAllVars(data: any) {
	PowerPoint.run(async context => {
		const slides = context.presentation.slides
		slides.load({ $all: true });
		await context.sync();

		let countSlides = slides.items.length;

		console.log("slides count:", countSlides);

		const itemText = (str: string) => {
			const div = document.createElement('div');
			div.innerHTML = str;
			return div.innerText
		}

		for (let i = 0; i < countSlides; i++) {
			const slide = context.presentation.slides.getItemAt(i);
			slide.shapes.load({ textFrame: { $all: true, textRange: { $all: true } } });
			await context.sync();

			slide.shapes.load({ textFrame: { textRange: { text: true } } });
			await context.sync();
			for (let v = 0; v < slide.shapes.items.length; v++) {
				// slide.shapes.items[v].load({ $all: true });
				// await context.sync();
				const var_data = slide.shapes.items[v].textFrame.textRange;
				// console.log(var_data.text);
				// console.log(slide.shapes.items[v].name);

				if (var_data.text.startsWith("{{")) {
					var x = data.product[data.variables.find((el: any) => el.name === var_data.text).column];
					var_data.text = itemText(x);
					context.sync();
				}
			}
		}

		// console.log(slide)
		
		// Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (res) => { // @ts-ignore
		// 	console.log(res.value.slides[0].index);
		// });

		// context.presentation.slides.load({ id: true, layout: {$all: true}, slideMaster: {$all: true} });
		// await context.sync();
	});
}
