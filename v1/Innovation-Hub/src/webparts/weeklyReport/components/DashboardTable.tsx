import * as React from "react";
import { useState, useEffect } from "react";

import "../ExternalRef/styleSheets/Styles.css";
import styles from "./WeeklyReport.module.scss";

let initialObj = {
  UserID: null,
  UserEmail: "",
  User: "",
  TWPH: null,
  TAH: null,
  TPH: null,
  ShowAll: false,
  data: {},
};
let maxDeliverableCount: number = null;
let resArr = [];
let html: string = "";

let excelInitialObj = {
  UserID: null,
  UserEmail: "",
  User: "",
  TWPH: null,
  TAH: null,
  TPH: null,
  ShowAll: false,
  data: {},
};
let excelMaxDeliverableCount: number = null;
let excelResArr = [];
let excelHtml: string = "";

interface projectDetails {
  ID: number;
  PBType: string;
  projectName: string;
  AH: number;
  PH: number;
  EndDate: string;
}
interface IHistoryData {
  FileLink: string;
  FileName: string;
  Sent: string;
  SentToName: string;

  FromName: string;
  ResponseDate: Date;
  Rating: number;
  Requests: string;
  Responses: string;
  ResponseComments: string;
  RequestComments: string;
  From: { ID: number; Name: string; Email: string };
  SentTo: { ID: number; Name: string; Email: string };
  PBID: number;
  PBType: string;
}
interface IData {
  ID: number;

  UserID: number;
  UserName: string;
  UserEmail: string;

  maxProjects: projectDetails[];

  BA: string;
  ActiveStatus: string;

  TH: number;
  AH: number;
  PH: number;

  ShowAll: boolean;

  Review: number;
  Edit: number;
  Assemble: number;
  SignOff: number;
  Publish: number;
  Returned: number;
  Feedback: number;
  Actioned: number;
  Endosed: number;
  SignedOff: number;
  RepeatedIssues: number;
  Major_Moderate: number;
  Incomplete: number;
  Quality: number;
  Style: number;

  ReviewData: IHistoryData[];
  EditData: IHistoryData[];
  AssembleData: IHistoryData[];
  SignOffData: IHistoryData[];
  PublishData: IHistoryData[];
  ReturnedData: IHistoryData[];
  FeedbackData: IHistoryData[];
  ActionedData: IHistoryData[];
  EndosedData: IHistoryData[];
  SignedOffData: IHistoryData[];
  RepeatedIssuesData: IHistoryData[];
  Major_ModerateData: IHistoryData[];
  IncompleteData: IHistoryData[];
  QualityData: IHistoryData[];
  StyleData: IHistoryData[];
}
interface IOnClickData {
  UserID: number | string;
  projectindex: number | string;
  Type: string;
}

interface IProps {
  masterData: IData[];
  filteredData: IData[];
  displayData: IData[];
  pageSwitch: any;
  getHtml: any;
}

const PHDashboardTable = (props: IProps) => {
  const dataManipulationFunction = () => {
    resArr = [];
    maxDeliverableCount = null;
    initialObj = {
      UserID: null,
      UserEmail: "",
      User: "",
      TWPH: null,
      TAH: null,
      TPH: null,
      ShowAll: false,
      data: {},
    };
    html = "";

    maxDeliverableCount = Math.max(
      ...props.displayData.map((o) => o.maxProjects.length)
    );

    //Dynamic Initial-Object Generation Starts

    for (let i = 1; i <= maxDeliverableCount; i++) {
      let tempStr = `{ "D${i}_R": "", "D${i}_E": "", "D${i}_A": "", "D${i}_S": "", "D${i}_P": "" ,"D${i}_PHAH": "","D${i}_PH": "","D${i}_AH": "" }`;

      initialObj = Object.assign(initialObj, JSON.parse(tempStr));
    }

    //Dynamic Initial-Object Generation Ends

    //Assigning value for Object Starts

    for (let i = 0; i < props.displayData.length; i++) {
      let obj = props.displayData[i];
      let sampleObj = { ...initialObj };

      sampleObj.UserID = obj.UserID;
      sampleObj.User = obj.UserName;
      sampleObj.UserEmail = obj.UserEmail;
      // sampleObj.TWPH = 0;
      sampleObj.TWPH = obj.TH.toString().match(/\./g)
        ? obj.TH.toFixed(2)
        : obj.TH;
      sampleObj.TAH = obj.AH.toString().match(/\./g)
        ? obj.AH.toFixed(2)
        : obj.AH;
      sampleObj.TPH = obj.PH.toString().match(/\./g)
        ? obj.PH.toFixed(2)
        : obj.PH;
      sampleObj.ShowAll = obj.ShowAll;
      sampleObj.data = { ...obj };

      for (let j = 0; j < obj.maxProjects.length; j++) {
        let project = obj.maxProjects[j];

        sampleObj[`D${j + 1}_R`] = countCalculator(project, obj, "ReviewData");
        sampleObj[`D${j + 1}_E`] = countCalculator(project, obj, "EditData");
        sampleObj[`D${j + 1}_A`] = countCalculator(
          project,
          obj,
          "AssembleData"
        );
        sampleObj[`D${j + 1}_S`] = countCalculator(project, obj, "SignOffData");
        sampleObj[`D${j + 1}_P`] = countCalculator(project, obj, "PublishData");
        sampleObj[`D${j + 1}_PHAH`] =
          project.AH || project.PH
            ? `<sup>${project.AH ? project.AH : 0}</sup>/
            <sub>${project.PH ? project.PH : 0}</sub>`
            : "";
        sampleObj[`D${j + 1}_PH`] = project.PH ? project.PH : "";
        sampleObj[`D${j + 1}_AH`] = project.AH ? project.AH : "";
        // sampleObj[`D${j + 1}_PHAH`] =
        //   project.AH || project.PH
        //     ? `<sup>${project.PH ? project.PH : 0}</sup>/
        //     <sub>${project.AH ? project.AH : 0}</sub>`
        //     : "";
      }

      resArr.push(sampleObj);
    }

    //Assigning value for Object Ends

    // HTML Binding Starts
    html = `<table class="ProductReportTable" cellspacing="0px">
    <thead>
    <tr id="thProductreport">
    <th scope="colgroup" rowspan="2" class=${styles.stickyHead} style="left:0;column-count:8;">User</th>
    <th title="Total weekly production hours" scope="colgroup" rowspan="2" class=${styles.stickyHead} style="left:363px;">TWPH</th>
    <th title="Total planned hours" scope="colgroup" rowspan="2" class=${styles.stickyHead} style="left:420px;padding-right:16px;">TPH</th>
    <th title="Total actual hours" scope="colgroup" rowspan="2" class=${styles.stickyHead} style="left:473px;">TAH</th>
    `;
    for (let i = 1; i <= maxDeliverableCount; i++) {
      i == 0
        ? (html += `<th colspan="6" scope="colgroup" class=${styles.stickyHead} style="left:517px;">D${i}</th>`)
        : (html += `<th colspan="6" scope="colgroup">D${i}</th>`);
    }

    html += `</tr><tr id="thSecondProductreport">`;
    for (let i = 1; i <= maxDeliverableCount; i++) {
      i == 0
        ? (html += `
      <th scope="col" class=${styles.stickyHead} style="left:517px;">R</th>
      <th scope="col" class=${styles.stickyHead} style="left:558px;">E</th>
      <th scope="col" class=${styles.stickyHead} style="left:600px;">A</th>
      <th scope="col" class=${styles.stickyHead} style="left:640px;">S</th>
      <th scope="col" class=${styles.stickyHead} style="left:682px;">P</th>
      <th scope="col" class=${styles.stickyHead} style="left:724px; column-count:2;"><strong><sup>AH</sup>/<sub>PH</sub></strong></th>
      `)
        : (html += `
      <th title="Review" scope="col">R</th>
      <th title="Edit" scope="col">E</th>
      <th title="Assemble" scope="col">A</th>
      <th title="Signoff" scope="col">S</th>
      <th title="Publish" scope="col">P</th>
      <th title="Actual hours/Planned hours" scope="col"><strong><sup>AH</sup>/<sub>PH</sub></strong></th>
      `);
    }

    html += `</tr></thead><tbody id="tbodyProductReport">`;

    resArr.forEach((arr) => {
      html += `<tr>
      <td class=${
        styles.stickyData
      } style="left:0;display: flex;align-items: center;" >
      <img style="border-radius: 50%;" src="/_layouts/15/userphoto.aspx?size=S&username=${
        arr.UserEmail
      }" alt="" width="30" height="30">
      <span style="margin-left: 10px;">${arr.User}</span>
      </td>
      <td class=${styles.stickyData} style="left:363px;"
        >${arr.TWPH}</td>
      <td class=${styles.stickyData} style="left:420px;background:
      ${arr.ShowAll == true ? "#f5e3e3 !important" : "#fff !important"}">${
        arr.TPH
      }</td>
      <td class=${styles.stickyData} style="left:473px;background:
      ${arr.ShowAll == true ? "#f5e3e3 !important" : "#fff !important"}">${
        arr.TAH
      }</td>`;

      for (let k = 1; k <= maxDeliverableCount; k++) {
        let valR: number = arr[`D${k}_R`]
          ? arr[`D${k}_R`].split("(")[1].replace(")", "")
          : 0;
        let valE: number = arr[`D${k}_E`]
          ? arr[`D${k}_E`].split("(")[1].replace(")", "")
          : 0;
        let valA: number = arr[`D${k}_A`]
          ? arr[`D${k}_A`].split("(")[1].replace(")", "")
          : 0;
        let valS: number = arr[`D${k}_S`]
          ? arr[`D${k}_S`].split("(")[1].replace(")", "")
          : 0;
        let valP: number = arr[`D${k}_P`]
          ? arr[`D${k}_P`].split("(")[1].replace(")", "")
          : 0;

        k == 0
          ? (html += `
        <td class="tdforonlick ${styles.stickyData}"
        style="left:517px;
        
        cursor:${arr[`D${k}_R`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_R`]}' 
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Review"}'>
        ${arr[`D${k}_R`]}</td>

        <td class="tdforonlick ${styles.stickyData}" 
        style="left:558px;
        cursor:${arr[`D${k}_E`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_E`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Edit"}'>
        ${arr[`D${k}_E`]}</td>

        <td class="tdforonlick ${styles.stickyData}" 
        style="left:600px;
        cursor:${arr[`D${k}_A`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_A`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Assemble"}'>
        ${arr[`D${k}_A`]}</td>

        <td class="tdforonlick ${styles.stickyData}" 
        style="left:640px;
        cursor:${arr[`D${k}_S`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_S`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"SignOff"}'>
        ${arr[`D${k}_S`]}</td>

        <td class="tdforonlick ${styles.stickyData}" 
        style="left:682px;
        cursor:${arr[`D${k}_P`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_P`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Publish"}'>
        ${arr[`D${k}_P`]}</td>

        <td class=${styles.stickyData} style="left:724px;">
        ${arr[`D${k}_PHAH`]}</td>
        `)
          : (html += `
        <td class="tdforonlick"
        style="left:517px; 
        background:
        ${valR >= 3 ? "#f5e3e3 !important" : "#fff !important"};
        cursor:${arr[`D${k}_R`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_R`]}' 
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Review"}'>
        ${arr[`D${k}_R`]}</td>

        <td class="tdforonlick" 
        style="left:558px; 
        background:
        ${valE >= 3 ? "#f5e3e3 !important" : "#fff !important"};
        cursor:${arr[`D${k}_E`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_E`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Edit"}'>
        ${arr[`D${k}_E`]}</td>

        <td class="tdforonlick" 
        style="left:600px; 
        background:
        ${valA >= 2 ? "#f5e3e3 !important" : "#fff !important"};
        cursor:${arr[`D${k}_A`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_A`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Assemble"}'>
        ${arr[`D${k}_A`]}</td>

        <td class="tdforonlick" 
        style="left:640px; 
        background:
        ${valS >= 2 ? "#f5e3e3 !important" : "#fff !important"};
        cursor:${arr[`D${k}_S`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_S`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"SignOff"}'>
        ${arr[`D${k}_S`]}</td>

        <td class="tdforonlick" 
        style="left:682px; 
        background:
        ${valP >= 2 ? "#f5e3e3 !important" : "#fff !important"};
        cursor:${arr[`D${k}_P`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_P`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Publish"}'>
        ${arr[`D${k}_P`]}</td>

        <td style="left:724px;">${arr[`D${k}_PHAH`]}</td>
        `);
      }
    });
    html += `</tr></tbody></table>`;
    // HTML Binding Ends
    document.getElementById("htmlTableBindID").innerHTML = html;

    var maxCount = document.getElementsByClassName("tdforonlick");

    for (let i = 0; i < maxCount.length; i++) {
      maxCount[i].addEventListener("click", (e) => {
        let _value = maxCount[i].getAttribute("data-value");
        let _data_obj = maxCount[i].getAttribute("data-object");
        if (_value) {
          pageSwitcher(JSON.parse(_data_obj));
        }
      });
    }
  };

  const dataManipulationFunctionforExport = () => {
    excelResArr = [];
    excelMaxDeliverableCount = null;
    excelInitialObj = {
      UserID: null,
      UserEmail: "",
      User: "",
      TWPH: null,
      TAH: null,
      TPH: null,
      ShowAll: false,
      data: {},
    };
    excelHtml = "";

    excelMaxDeliverableCount = Math.max(
      ...props.filteredData.map((o) => o.maxProjects.length)
    );

    //Dynamic Initial-Object Generation Starts

    for (let i = 1; i <= excelMaxDeliverableCount; i++) {
      let tempStr = `{ "D${i}_R": "", "D${i}_E": "", "D${i}_A": "", "D${i}_S": "", "D${i}_P": "" ,"D${i}_PHAH": "","D${i}_PH": "","D${i}_AH": "" }`;

      excelInitialObj = Object.assign(excelInitialObj, JSON.parse(tempStr));
    }

    //Dynamic Initial-Object Generation Ends

    //Assigning value for Object Starts

    for (let i = 0; i < props.filteredData.length; i++) {
      let obj = props.filteredData[i];
      let sampleObj = { ...excelInitialObj };

      sampleObj.UserID = obj.UserID;
      sampleObj.User = obj.UserName;
      sampleObj.UserEmail = obj.UserEmail;
      // sampleObj.TWPH = 0;
      sampleObj.TWPH = obj.TH.toString().match(/\./g)
        ? obj.TH.toFixed(2)
        : obj.TH;
      sampleObj.TAH = obj.AH.toString().match(/\./g)
        ? obj.AH.toFixed(2)
        : obj.AH;
      sampleObj.TPH = obj.PH.toString().match(/\./g)
        ? obj.PH.toFixed(2)
        : obj.PH;
      sampleObj.data = { ...obj };

      for (let j = 0; j < obj.maxProjects.length; j++) {
        let project = obj.maxProjects[j];

        sampleObj[`D${j + 1}_R`] = countCalculator(project, obj, "ReviewData");
        sampleObj[`D${j + 1}_E`] = countCalculator(project, obj, "EditData");
        sampleObj[`D${j + 1}_A`] = countCalculator(
          project,
          obj,
          "AssembleData"
        );
        sampleObj[`D${j + 1}_S`] = countCalculator(project, obj, "SignOffData");
        sampleObj[`D${j + 1}_P`] = countCalculator(project, obj, "PublishData");
        sampleObj[`D${j + 1}_PHAH`] =
          project.AH || project.PH
            ? `<sup>${project.AH ? project.AH : 0}</sup>/
            <sub>${project.PH ? project.PH : 0}</sub>`
            : "";
        sampleObj[`D${j + 1}_PH`] = project.PH ? project.PH : "";
        sampleObj[`D${j + 1}_AH`] = project.AH ? project.AH : "";
        // sampleObj[`D${j + 1}_PHAH`] =
        //   project.AH || project.PH
        //     ? `<sup>${project.PH ? project.PH : 0}</sup>/
        //     <sub>${project.AH ? project.AH : 0}</sub>`
        //     : "";
      }

      excelResArr.push(sampleObj);
    }

    //Assigning value for Object Ends

    // excel Html
    excelHtml = `<table class="ProductReportTable" cellspacing="0px">
    <thead>
    <tr id="thProductreport">
    <th scope="colgroup" rowspan="2" class=${styles.stickyHead} style="left:0;column-count:8;">User</th>
    <th title="Total weekly production hours" scope="colgroup" rowspan="2" class=${styles.stickyHead} style="left:363px;">TWPH</th>
    <th title="Total planned hours" scope="colgroup" rowspan="2" class=${styles.stickyHead} style="left:420px;padding-right:16px;">TPH</th>
    <th title="Total actual hours" scope="colgroup" rowspan="2" class=${styles.stickyHead} style="left:473px;">TAH</th>
    `;
    for (let i = 1; i <= excelMaxDeliverableCount; i++) {
      i == 0
        ? (excelHtml += `<th colspan="7" scope="colgroup" class=${styles.stickyHead} style="left:517px;">D${i}</th>`)
        : (excelHtml += `<th colspan="7" scope="colgroup">D${i}</th>`);
    }

    excelHtml += `</tr><tr id="thSecondProductreport">`;
    for (let i = 1; i <= excelMaxDeliverableCount; i++) {
      i == 0
        ? (excelHtml += `
      <th scope="col" class=${styles.stickyHead} style="left:517px;">R</th>
      <th scope="col" class=${styles.stickyHead} style="left:558px;">E</th>
      <th scope="col" class=${styles.stickyHead} style="left:600px;">A</th>
      <th scope="col" class=${styles.stickyHead} style="left:640px;">S</th>
      <th scope="col" class=${styles.stickyHead} style="left:682px;">P</th>
      <th scope="col" class=${styles.stickyHead} style="left:724px;">AH</th>
      <th scope="col" class=${styles.stickyHead} style="left:724px;">PH</th>
      `)
        : (excelHtml += `
      <th title="Review" scope="col">R</th>
      <th title="Edit" scope="col">E</th>
      <th title="Assemble" scope="col">A</th>
      <th title="Signoff" scope="col">S</th>
      <th title="Publish" scope="col">P</th>
      <th title="AH" scope="col">AH</th>
      <th title="PH" scope="col">PH</th>
      `);
    }

    excelHtml += `</tr></thead><tbody id="tbodyProductReport">`;

    excelResArr.forEach((arr) => {
      excelHtml += `<tr>
      <td class=${styles.stickyData} style="left:0;" >
      <span>${arr.User}</span>
      </td>
      <td class=${styles.stickyData} style="left:363px;">${arr.TWPH}</td>
      <td class=${styles.stickyData} style="left:420px;">${arr.TPH}</td>
      <td class=${styles.stickyData} style="left:473px;">${arr.TAH}</td>`;

      for (let k = 1; k <= excelMaxDeliverableCount; k++) {
        k == 0
          ? (excelHtml += `
        <td class="tdforonlick ${styles.stickyData}"
        style="left:517px;
        cursor:${arr[`D${k}_R`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_R`]}' 
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Review"}'>
        ${arr[`D${k}_R`]}</td>

        <td class="tdforonlick ${styles.stickyData}" 
        style="left:558px;
        cursor:${arr[`D${k}_E`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_E`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Edit"}'>
        ${arr[`D${k}_E`]}</td>

        <td class="tdforonlick ${styles.stickyData}" 
        style="left:600px;
        cursor:${arr[`D${k}_A`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_A`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Assemble"}'>
        ${arr[`D${k}_A`]}</td>

        <td class="tdforonlick ${styles.stickyData}" 
        style="left:640px;
        cursor:${arr[`D${k}_S`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_S`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"SignOff"}'>
        ${arr[`D${k}_S`]}</td>

        <td class="tdforonlick ${styles.stickyData}" 
        style="left:682px;
        cursor:${arr[`D${k}_P`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_P`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Publish"}'>
        ${arr[`D${k}_P`]}</td>

        <td class=${styles.stickyData} style="left:724px;">
        ${arr[`D${k}_AH`]}</td>
        <td class=${styles.stickyData} style="left:724px;">
        ${arr[`D${k}_PH`]}</td>
        `)
          : (excelHtml += `
        <td class="tdforonlick"
        style="left:517px; cursor:${arr[`D${k}_R`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_R`]}' 
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Review"}'>
        ${arr[`D${k}_R`]}</td>

        <td class="tdforonlick" 
        style="left:558px; cursor:${arr[`D${k}_E`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_E`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Edit"}'>
        ${arr[`D${k}_E`]}</td>

        <td class="tdforonlick" 
        style="left:600px; cursor:${arr[`D${k}_A`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_A`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Assemble"}'>
        ${arr[`D${k}_A`]}</td>

        <td class="tdforonlick" 
        style="left:640px; cursor:${arr[`D${k}_S`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_S`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"SignOff"}'>
        ${arr[`D${k}_S`]}</td>

        <td class="tdforonlick" 
        style="left:682px; cursor:${arr[`D${k}_P`] ? "pointer" : "default"};"
        data-value='${arr[`D${k}_P`]}'
        data-object='{"UserID":"${arr.UserID}",
        "projectindex":${k - 1},"Type":"Publish"}'>
        ${arr[`D${k}_P`]}</td>

        <td style="left:724px;">${arr[`D${k}_AH`]}</td>
        <td style="left:724px;">${arr[`D${k}_PH`]}</td>
        `);
      }
    });
    excelHtml += `</tr></tbody></table>`;

    props.getHtml(excelHtml);
  };

  const countCalculator = (project, data, funcType: string): string => {
    let filteredArr = [];
    let tempData = [];

    tempData = data[funcType].filter((obj) => {
      return obj.PBType;
    });

    filteredArr = tempData.filter((_data) => {
      let firstPBID: number = parseInt(_data.PBID.replace(",", ""));
      let secondPBID: number = parseInt(project.ID);

      return firstPBID == secondPBID && _data.PBType == project.PBType;
    });

    let unique = filteredArr
      .map((item) => item.FileName)
      .filter((value, index, self) => self.indexOf(value) === index);

    return `${unique.length}(${filteredArr.length})`;
  };
  const pageSwitcher = (_obj: IOnClickData) => {
    let filteredArr: IData[] = props.masterData.filter((_data: IData) => {
      return _data.UserID == _obj.UserID;
    });

    let maxProjects = filteredArr[0].maxProjects;
    let targetProject: projectDetails =
      maxProjects.length > 0 ? maxProjects[_obj.projectindex] : null;

    let finalArr: IHistoryData[] =
      targetProject != null
        ? filteredArr[0][`${_obj.Type}Data`].filter((_arr: IHistoryData) => {
            let newPBID: number = _arr.PBID
              ? parseInt(_arr.PBID.toString().replace(",", ""))
              : null;
            return newPBID == targetProject.ID;
          })
        : [];

    props.pageSwitch(
      true,
      filteredArr[0].UserName,
      filteredArr[0].UserEmail,
      `${_obj.Type},${targetProject.projectName},${targetProject.EndDate}`,
      finalArr
    );
  };

  useEffect(() => {
    dataManipulationFunction();
    dataManipulationFunctionforExport();
  }, [props.displayData]);

  useEffect(() => {}, []);
  return (
    <div>
      <div
        id="htmlTableBindID"
        className={styles.tableWrapper}
        style={{ overflowX: "auto" }}
      ></div>
    </div>
  );
};

export default PHDashboardTable;
