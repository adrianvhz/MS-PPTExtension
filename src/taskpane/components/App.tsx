import * as React from "react";
import { PrimaryButton as Button, DefaultButton } from "@fluentui/react";
import Header from "./Header";
import Footer from "./Footer";
import List, { ListItem } from "./List";
import Progress from "./Progress";
import NotificationBar from "./NotificationBar";
import RequestedButtons from "./RequestedButtons";
import ReplaceVars from "./ReplaceVars";

/* global console, Office, require */

export interface AppProps {
	title: string;
	isOfficeInitialized: boolean;
}

export interface AppState {
	listItems: ListItem[];
	notification: {
		message: string;
		error: boolean;
	}
}

export default function App(props: any) {
	const [isMasterTemplate, setIsMasterTemplate] = React.useState(null);
	const [filename, setFilename] = React.useState("");
	const [notification, setNotification] = React.useState({
		message: "none",
		error: false
	});

	const { title, isOfficeInitialized } = props;
	
	console.log("is master template:", isMasterTemplate);

	const handleFilename = (name: string) => {
		setFilename(name);
	}

	React.useEffect(() => {
		// console.log(Office.context.document.getFilePropertiesAsync(null, (x) => {
		// 	console.log(x.value);
		// }));
		getIsMasterTemplate(handleFilename).then((res) => {
			setIsMasterTemplate(res);
		});
		console.log("Permisos de archivo:", Office.context.document.mode);	
	}, []);

	if (!isOfficeInitialized) {
		return (
			<Progress
				title={title}
				logo={require("./../../../assets/logo-filled.png")}
				message="Please sideload your addin to see app body."
			/>
		)
	}

	if (isMasterTemplate === null) {
		return <div>NONE</div>
	}

	return (
		<div className="ms-welcome" style={{ padding: "1rem 1.5rem" }}>
			{/* <Header logo={require("./../../../assets/logo-filled.png")} title={title} message="DEMO" /> */}

			{/* <h3 style={{ marginTop: 0, marginBottom: "1rem" }}>
				Producto: {"{prod_name}"}
			</h3> */}
			{isMasterTemplate
				? <RequestedButtons />
				: <ReplaceVars filename={filename} />
			}
			
			{/* <List message="demo">
				<Button
					className="ms-welcome__action"
					iconProps={{ iconName: "ChangeEntitlements" }}
					onClick={() => replaceVars(setNotification)}
				>
					Replace Vars
				</Button>
				<Button onClick={insertImg} iconProps={{ iconName: "MediaAdd" }}>
					Insert Image
				</Button>
				<Button
					className="ms-welcome__action"
					iconProps={{ iconName: "AddToShoppingList" }}
					onClick={newDefaultSlide}
					style={{ backgroundColor: "#6ba66b", border: "1px solid #6ba66b" }}
				>
					New Slide With Default Vars
				</Button>
				<DefaultButton
					className="ms-welcome__action"
					iconProps={{ iconName: "ChevronRight" }}
				>
					Add Shape
				</DefaultButton>
			</List> */}
			<NotificationBar notification={notification} />
			{/* <Footer logo="a" message="a" title="s" /> */}
		</div>
	)
}

const getImageAsBase64String = () => {
	// A production add-in code could get an image from an
	// online source and pass it to a library function that
	// converts to base 64.
	return "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"
}

const insertImg = async () => {
	Office.context.document.setSelectedDataAsync(getImageAsBase64String(), {
		coercionType: Office.CoercionType.Image,
		imageLeft: 50,
		imageTop: 50,
		imageWidth: 400
	},
	function (asyncResult) {
		if (asyncResult.status === Office.AsyncResultStatus.Failed) {
			this.setState({
				notification: {
					message: asyncResult.error.message,
					error: true
				}
			})
			console.log(asyncResult.error.message);
		}
	});
}

const replaceVars = async (handleNotification: Function) => {
	const slideIndex = document.getElementById("input-000") as HTMLInputElement;
	const title = document.getElementById("input-001") as HTMLInputElement;
	const subtitle = document.getElementById("input-002") as HTMLInputElement;

	var currentSlide = 0;
	
	await PowerPoint.run(async (context) => {
		// get current slide index
		await new Promise((resolve, reject) => {
			Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (asyncResult) => { // @ts-ignore
				if (asyncResult.status == "failed") {
					console.log(asyncResult.error.message)
					reject(null);
				}
				else { // @ts-ignore
					currentSlide = asyncResult.value.slides[0].index;
					resolve(currentSlide);
				}
			})
		});

		var slide = context.presentation.slides.getItemAt(+slideIndex.value || currentSlide-1);
		slide.shapes.load({
			textFrame: {
				$all: true,
				textRange: {
					$all: true
				}
			}
		});

		await context.sync();
		
		for (var i = 0; i < slide.shapes.items.length; i++) {
			if (slide.shapes.items[i].textFrame.textRange.text === "{{title}}") {
				slide.shapes.items[i].textFrame.textRange.text = title.value || "My Custom Title";
				console.log("Successfully replaced title!");
				handleNotification({
					message: "Successfully replaced!",
					error: false
				});

				await context.sync();
			} else {
				handleNotification({
					message: "{{title}} variable not found",
					error: true
				});
			}

			if (slide.shapes.items[i].textFrame.textRange.text === "{{subtitle}}") {
				slide.shapes.items[i].textFrame.textRange.text = subtitle.value || "My Custom Subtitle";
				console.log("Successfully replaced subtitle!");
				handleNotification({
					message: "Successfully replaced!",
					error: false
				});

				await context.sync();
			} else {
				handleNotification({
					message: "{{subtitle}} variable not found",
					error: true
				});
			}
		}
	})
}

const newDefaultSlide = async () => {
	await PowerPoint.run(async (context) => {
		context.presentation.slides.add();
		context.presentation.slides.load({ $all: true });
		await context.sync();

		const newSlideIndex = context.presentation.slides.items.length;
		const newSlide = context.presentation.slides.getItemAt(newSlideIndex);

		newSlide.shapes.load({
			$all: true,
			textFrame: {
				$all: true
			}
		});
		await context.sync();

		const shape1 = newSlide.shapes.getItemAt(0);
		const shape2 = newSlide.shapes.getItemAt(1);

		shape1.textFrame.textRange.text = "{{title}}";
		await context.sync();
		shape2.textFrame.textRange.text = "{{subtitle}}";
	})
}

// class AppClass extends React.Component<AppProps, AppState> {
// 	constructor(props: any, context: any) {
// 		super(props, context);
// 		this.state = {
// 			listItems: [],
// 			notification: {
// 				message: "none",
// 				error: false
// 			}
// 		}
// 	}

// 	componentDidMount() {
// 		this.setState({
// 			listItems: [
// 				{
// 					icon: "Design",
// 					id: 0,
// 					name: "Slide_Index",
// 					info: "(Default: slide actual)"
// 				},
// 				{
// 					// icon: "Ribbon",
// 					id: 1,
// 					name: "Title",
// 					info: ""
// 				},
// 				{
// 					// icon: "Unlock",
// 					id: 2,
// 					name: "Subtitle",
// 					info: ""
// 				}
// 			]
// 		});
// 	}

// 	getImageAsBase64String = () => {
// 		// A production add-in code could get an image from an
// 		// online source and pass it to a library function that
// 		// converts to base 64.
// 		return "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"
// 	}

// 	insertImg = async () => {
// 		Office.context.document.setSelectedDataAsync(this.getImageAsBase64String(), {
// 			coercionType: Office.CoercionType.Image,
// 			imageLeft: 50,
// 			imageTop: 50,
// 			imageWidth: 400
// 		},
// 		function (asyncResult) {
// 			if (asyncResult.status === Office.AsyncResultStatus.Failed) {
// 				this.setState({
// 					notification: {
// 					message: asyncResult.error.message,
// 					error: true
// 					}
// 				})
// 				console.log(asyncResult.error.message);
// 			}
// 		});
// 	}

// 	// addShape = async () => {
// 	//   await PowerPoint.run(async function(context) {
// 	//     var shapes = context.presentation.slides.getItemAt(0).shapes;
// 	//     var rectangle = shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
// 	//     rectangle.left = 100;
// 	//     rectangle.top = 100;
// 	//     rectangle.height = 150;
// 	//     rectangle.width = 150;
// 	//     rectangle.name = "Square";
// 	//     await context.sync();
// 	//   })
// 	// }

// 	insertText = async () => {
// 		Office.context.document.setSelectedDataAsync(
// 			"Test Text",
// 			{
// 				coercionType: Office.CoercionType.Text
// 			},
// 			(result) => {
// 				if (result.status === Office.AsyncResultStatus.Failed) {
// 					this.setState({
// 					notification: {
// 						message: result.error.message,
// 						error: true
// 					}
// 					})
// 					console.error(result.error.message);
// 				} else {
// 					this.setState({
// 					notification: {
// 						message: "Text inserted successfully!",
// 						error: false
// 					}
// 					})
// 				}
// 			}
// 		);
// 	}

// 	replaceVars = async () => {
// 		const slideIndex = document.getElementById("Slide_Index") as HTMLInputElement;
// 		const title = document.getElementById("Title") as HTMLInputElement;
// 		const subtitle = document.getElementById("Subtitle") as HTMLInputElement;

// 		const body = new URLSearchParams();
// 		body.append("id", "8240");

// 		const x = await fetch("https://wserv-qa.proeducative.com/program/getProducDetailById", {
// 			method: "POST",
// 			headers: {
// 				"Content-Type": "application/x-www-form-urlencoded",
// 				"Authorization": "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJpYXQiOjE2NzM0NzUwNDgsImV4cCI6MTY3MzUwNTA0OCwia2V5IjoiQUgzaUgxNlBYUzIifQ.-vdeFay-epMdavJPvL7Le07TkVCSjESFONXURVfuYZA"
// 			},
// 			body: new URLSearchParams([["id", "8240"]])
// 		}).then(r => r.json());

// 		console.log(x)

// 		var currentSlide = 0;
		
// 		await PowerPoint.run(async (context) => {
// 			// get current slide index
// 			await new Promise((resolve, reject) => {
// 				Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (asyncResult) => { // @ts-ignore
// 					if (asyncResult.status == "failed") {
// 						console.log(asyncResult.error.message)
// 						reject(null);
// 					}
// 					else { // @ts-ignore
// 						currentSlide = asyncResult.value.slides[0].index;
// 						resolve(currentSlide);
// 					}
// 				})
// 			});

// 			var slide = context.presentation.slides.getItemAt(+slideIndex.value || currentSlide-1);
// 			slide.shapes.load({
// 				textFrame: {
// 					$all: true,
// 					textRange: {
// 						$all: true
// 					}
// 				}
// 			});

// 			await context.sync();
			
// 			for (var i = 0; i < slide.shapes.items.length; i++) {
// 				if (slide.shapes.items[i].textFrame.textRange.text === "{{title}}") {
// 					slide.shapes.items[i].textFrame.textRange.text = title.value || JSON.stringify(x.calculatedCosts[0]) || "My Custom Title";
// 					console.log("Successfully replaced title!");
// 					this.setState({
// 						notification: {
// 							message: "Successfully replaced!",
// 							error: false
// 						}
// 					});

// 					await context.sync();
// 				} else {
// 					this.setState({
// 						notification: {
// 							message: "{{title}} variable not found!",
// 							error: true
// 						}
// 					});
// 				}

// 				if (slide.shapes.items[i].textFrame.textRange.text === "{{subtitle}}") {
// 					slide.shapes.items[i].textFrame.textRange.text = subtitle.value || "My Custom Subtitle";
// 					console.log("Successfully replaced subtitle!");
// 					this.setState({
// 						notification: {
// 							message: "Successfully replaced!",
// 							error: false
// 						}
// 					});

// 					await context.sync();
// 				} else {
// 					this.setState({
// 						notification: {
// 							message: "{{subtitle}} variable not found!",
// 							error: true
// 						}
// 					});
// 				}
// 			}
// 		})
// 	}

// 	newDefaultSlide = async () => {
// 		await PowerPoint.run(async (context) => {
// 			context.presentation.slides.add();

// 			context.presentation.slides.load({ $all: true });
// 			await context.sync();

// 			const newSlideIndex = context.presentation.slides.items.length;
// 			const newSlide = context.presentation.slides.getItemAt(newSlideIndex);

// 			newSlide.shapes.load({
// 				$all: true,
// 				textFrame: {
// 					$all: true
// 				}
// 			});
// 			await context.sync();

// 			const shape1 = newSlide.shapes.getItemAt(0);
// 			const shape2 = newSlide.shapes.getItemAt(1);

// 			shape1.textFrame.textRange.text = "{{title}}";
// 			await context.sync();
// 			shape2.textFrame.textRange.text = "{{subtitle}}";
// 		})
// 	}

// 	render() {
// 		const { title, isOfficeInitialized } = this.props;

// 		if (!isOfficeInitialized) {
// 			return (
// 				<Progress
// 					title={title}
// 					logo={require("./../../../assets/logo-filled.png")}
// 					message="Please sideload your addin to see app body."
// 				/>
// 			)
// 		}

// 		return (
// 			<div className="ms-welcome">
// 				{/* <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="DEMO" /> */}
// 				<List message="Demo" items={this.state.listItems}>
// 					<Button className="ms-welcome__action" iconProps={{ iconName: "ChangeEntitlements" }} onClick={this.replaceVars}>
// 						Replace Vars
// 					</Button>
// 					<Button className="ms-welcome__action" iconProps={{ iconName: "InsertSignatureLine" }} onClick={this.insertText} style={{margin: ".3rem 0"}}>
// 						Insert Text Selection
// 					</Button>
// 					<Button  onClick={this.insertImg}>
// 						Insert Image
// 					</Button>
// 					<Button className="ms-welcome__action" iconProps={{ iconName: "ImageDiff" }} onClick={this.newDefaultSlide} style={{backgroundColor: "#6ba66b", border: "1px solid #6ba66b", marginTop: ".3rem"}}>
// 						New Slide With Default Vars
// 					</Button>
// 					<DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }}>
// 						Add Shape
// 					</DefaultButton>
// 					{}
// 				</List>
// 				<NotificationBar notification={this.state.notification} />
// 			</div>
// 		)
// 	}
// }





async function getIsMasterTemplate(handleFilename) {
	let filename = await getFileName(handleFilename); 

	if (filename.startsWith("MA")) return true;
	else if (filename.startsWith("PA")) return false;
	else return null;
}

const getFileName: (handleFilename) => Promise<string> = async (handleFilename) => {
	return new Promise((resolve) => {
		Office.context.document.getFilePropertiesAsync(null, (res) => {
			if (res && res.value && res.value.url) {
				let name = res.value.url.slice(res.value.url.lastIndexOf("/") + 1);
				handleFilename(name);
				resolve(name);
			} else {
				resolve("");
			}
		});
	});
}
