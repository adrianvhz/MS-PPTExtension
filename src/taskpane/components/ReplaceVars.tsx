import * as React from "react";

export default function ReplaceVars({ filename }) {
	console.log("filename", filename)
	React.useEffect(() => {
		fetch(`https://idcloudsystem.com/api/products/${filename}/detail`, {
			method: "POST",
			headers: {
				"Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpYXQiOjE2NzQwNTk3MTYsImV4cCI6MTY3NDE0NjExNiwia2V5IjoidXNlcnBpZWNlIn0.AX6MoJVyeG54noHdy3GglwsXf-aY7J8TijxfPosIeWM"
			},
		})
			.then(raw => raw.json()).then(res => {
				console.log(res);
				replaceAllVars(res.data);
			})
	}, []);

	return (
		<div>Reemplazando...</div>
	)
}

async function replaceAllVars(data: any) {
	PowerPoint.run(async context => {
		const slides = context.presentation.slides
		slides.load({ $all: true });
		await context.sync();

		let countSlides = slides.items.length;

		console.log("slides count:", countSlides);

		for (let i = 0; i < countSlides; i++) {
			const slide = context.presentation.slides.getItemAt(i);
			slide.shapes.load({ textFrame: { $all: true, textRange: { $all: true } } });
			await context.sync();

			slide.shapes.load({ textFrame: { textRange: { text: true } }, name: true });
			await context.sync();
			for (let v = 0; v < slide.shapes.items.length; v++) {
				// slide.shapes.items[v].load({ $all: true });
				// await context.sync();
				const var_data = slide.shapes.items[v].textFrame.textRange;
				console.log(var_data.text);
				console.log(slide.shapes.items[v].name);
				console.log(var_data.text === "{{PRESENTATION_PRODUCTO}}");
				if (var_data.text.startsWith("{{")) {
					// var_data.text = data.product.presentation;
					var_data.text = data.product[var_data.text.slice(2).split("_")[0].toLowerCase()];
					context.sync();
				} else {
					continue;
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
