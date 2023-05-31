const express = require("express");
const morgan = require("morgan");
const app = express();

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(morgan("dev"));
app.use(express.static("public"));

app.use("/", require("./routes/main"));

const port = process.env.PORT || 3000;
app.listen(port, (req, res) => {
  console.log(`Server is running on ${port}`);
});
