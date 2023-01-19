import sign from "jwt-encode";

export default function generateToken() {
	// const token = sign({ key: "hVmYq3t6" }, "v9y$B&E)H@Mc", { expiresIn: "2d" });
	console.log("aca", 1, process.env.SECRET_KEY)
	const token = sign({ key: "hVmYq3t6" }, process.env.SECRET_KEY);
	window.localStorage.setItem("tk", token);
	return token;
}
