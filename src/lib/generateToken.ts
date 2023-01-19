import jtw from "jsonwebtoken";

export default function generateToken() {
	const token = jtw.sign({ key: "hVmYq3t6" }, "v9y$B&E)H@Mc", { expiresIn: "2d" });
	window.localStorage.setItem("tk", token);
	return token;
}
