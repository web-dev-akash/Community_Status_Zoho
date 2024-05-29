const express = require("express");
const cors = require("cors");
const pLimit = require("p-limit");
const axios = require("axios");
const path = require("path");
const xlsx = require("xlsx");
const multer = require("multer");
const fs = require("fs");
const { promisify } = require("util");
const unlinkAsync = promisify(fs.unlink);
require("dotenv").config();
const limit = pLimit(20);
const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(cors());
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REFRESH_TOKEN = process.env.REFRESH_TOKEN;
const PORT = process.env.PORT || 8080;

const storage = multer.diskStorage({
  destination: path.join(__dirname, "uploads"),
  filename: function (req, file, cb) {
    cb(null, file.originalname);
  },
});
const upload = multer({ storage: storage });

app.use(express.static(path.join(__dirname, "template")));

app.get("/", (req, res) => {
  res.sendFile(`index.html`);
});

const getZohoToken = async () => {
  try {
    const res = await axios.post(
      `https://accounts.zoho.com/oauth/v2/token?client_id=${CLIENT_ID}&grant_type=refresh_token&client_secret=${CLIENT_SECRET}&refresh_token=${REFRESH_TOKEN}`
    );
    console.log(res.data);
    const token = res.data.access_token;
    return token;
  } catch (error) {
    res.send({
      error,
    });
  }
};

const updateContactOnZoho = async ({ phone, config, group }) => {
  if (!Number(phone)) {
    return { phone, message: "Not a phone number" };
  }
  const contact = await axios.get(
    `https://www.zohoapis.com/crm/v2/Contacts/search?phone=${phone}`,
    config
  );
  if (!contact || !contact.data || !contact.data.data) {
    return { phone, message: "No Contact Found" };
  }
  // const key = "Premium_Group";
  const key = "Joined_Wisechampions";
  const contactId = contact.data.data[0].id;
  const alreadyJoined = contact.data.data[0][key];
  console.log(alreadyJoined);
  if (alreadyJoined) {
    return { phone, message: "Already in Community" };
  }
  const date = new Date();
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, "0");
  const day = date.getDate().toString().padStart(2, "0");
  const formattedDate = `${year}-${month}-${day}`;
  const body = {
    data: [
      {
        id: contactId,
        [key]: formattedDate,
        $append_values: {
          [key]: true,
        },
      },
    ],
    duplicate_check_fields: ["id"],
    apply_feature_execution: [
      {
        name: "layout_rules",
      },
    ],
    trigger: ["workflow"],
  };
  await axios.post(
    `https://www.zohoapis.com/crm/v3/Contacts/upsert`,
    body,
    config
  );
  return { phone, message: "Success" };
};

app.post("/view", upload.array("file", 50), async (req, res) => {
  const files = req.files;
  console.log(files);
  if (files.length === 0) {
    return res
      .status(400)
      .send(
        `<h1 style="display:grid;place-items:center;min-height:100vh;">No files were uploaded.</h1>`
      );
  }
  try {
    const currentUsers = [];
    for (const file of files) {
      console.log(file.path);
      const workbook = xlsx.readFile(file.path);
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const data = xlsx.utils.sheet_to_json(sheet);
      for (let i = 0; i < data.length; i++) {
        const phone = data[i]["Phone Numbers"]?.toString().replace(/[ +]/g, "");
        const groupName = data[i]["Group Membership"];
        currentUsers.push({ phone, groupName });
      }
      await unlinkAsync(file.path);
    }
    const token = await getZohoToken();
    const config = {
      headers: {
        Authorization: `Zoho-oauthtoken ${token}`,
        "Content-Type": "application/json",
      },
    };
    const result = await Promise.all(
      currentUsers.map(async (user) => {
        const [data] = await Promise.all([
          limit(() =>
            updateContactOnZoho({
              phone: user.phone,
              group: user.groupName,
              config,
            })
          ),
        ]);
        return data;
      })
    );
    return res.send({ data: result });
  } catch (error) {
    console.error("Error reading Excel file:", error);
    res.status(500).send({ error: "Error reading Excel file." });
    for (const file of files) {
      await unlinkAsync(file.path);
    }
    return;
  }
});

app.listen(PORT, () => {
  console.log(`http://localhost:${PORT}`);
});
