import sign from "jwt-encode";

export default function generateToken() {
	// const token = sign({ key: "hVmYq3t6" }, "v9y$B&E)H@Mc", { expiresIn: "2d" });
	const token = sign({ key: "hVmYq3t6" }, process.env.SECRET_KEY);
	return token;
}
