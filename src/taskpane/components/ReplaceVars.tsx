import * as React from "react";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import generateToken from "../../lib/generateToken";

export default function ReplaceVars({ filename }) {
	const [isLoading, setIsLoading] = React.useState(false);

	React.useEffect(() => {
		const alias = filename.split("-")[0];
		
		setIsLoading(true);
		fetch(`https://idcloudsystem.com/api/products/${alias}/detail`, {
			method: "POST",
			headers: {
				"Authorization": "Bearer " + generateToken()
			},
		})
		.then(raw => raw.json()).then(res => {
				console.log(res)
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
				const var_data = slide.shapes.items[v].textFrame.textRange;

				if (var_data.text.startsWith("{{")) {
					var x = data.product[data.variables.find((el: any) => el.name === var_data.text).column];
					var_data.text = itemText(x);
				}
			}
			context.sync();
		}
	});
}
