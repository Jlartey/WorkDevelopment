const mongoose = require("mongoose");
const User = require("./User");
// mongoose.connect(
//   "mongodb://localhost/testdb/",
//   () => {
//     console.log("connected");
//   },
//   (e) => console.error(e)
// );

mongoose
  .connect("mongodb://localhost/testdb")
  .then(() => {
    console.log("Connected to MongoDB");
  })
  .catch((err) => {
    console.error("Error connecting to MongoDB:", err);
  });

run();
async function run() {
  User.create({});
  const user = new User({ c });
  await user.save();
  console.log(user);
}
