const { text } = require("express");
var express = require("express");
var router = express.Router();
const fs = require("fs");
var xlsx = require("xlsx");

/* GET home page. */
router.get("/", function (req, res, next) {
 

  
  const file = xlsx.readFile("Filename.xlsx");
  let data = []
  const sheets = file.SheetNames;
  for (let i = 0; i < sheets.length; i++) {
    const temp = xlsx.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
    temp.forEach((res) => {
      data.push(res);
    });
  }
  console.log(data[0].name);



  res.status(200).render("index", { data: data });
  // res.render('index', { title: 'Express' });
});
router.post("/", function (req, res, next) {
  name = req.body.name;
  link = req.body.link;
  console.log(name);
  console.log(link);

  

  var jsn = [
    {
      name: req.body.name,
      link: req.body.link,
    },
  ];

  var data = "";

  for (var i = 0; i < jsn.length; i++) {
    data = data + jsn[i].name + "\t" + jsn[i].link + "\n";
  }
  fs.appendFile("Filename.xlsx", data, (err) => {
    if (err) throw err;
    console.log("File created");
  });

  res.redirect("/");
});

router.get("/index1.ejs", function (req, res, next) {
  const file = xlsx.readFile("Filename.xlsx");
  let data = []
  const sheets = file.SheetNames;
  for (let i = 0; i < sheets.length; i++) {
    const temp = xlsx.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
    temp.forEach((res) => {
      data.push(res);
    });
  }
  var jsn={
    name:"",
    text:"",
    link:"",
    thanks:"",
    date:"",
   }
  console.log(data[0].name);
  res.render('index1.ejs',{data:data,details:jsn});
});
router.post("/index1.ejs/output", function (req, res, next) {
  const file = xlsx.readFile("Filename.xlsx");
  let data = []
  const sheets = file.SheetNames;
  for (let i = 0; i < sheets.length; i++) {
    const temp = xlsx.utils.sheet_to_json(file.Sheets[file.SheetNames[i]]);
    temp.forEach((res) => {
      data.push(res);
    });
  }

console.log(req.body.type)
 console.log(req.body.thanks)
 console.log(req.body.first_name)
 
 var jsn={
  name:req.body.first_name,
  text:req.body.textarea,
  link:req.body.type,
  thanks:req.body.thanks,
  
  date:req.body.date,
  greet:"Dear",
  option:"Your selected option is"
 }
 

res.render("index1.ejs",{details:jsn,data:data})
});
 



module.exports = router;
