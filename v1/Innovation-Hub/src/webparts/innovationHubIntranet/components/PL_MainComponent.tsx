import * as React from "react";
import { useState, useEffect } from "react";
import "@pnp/sp/webs";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";

import PL_BusinessArea from "./PL_BusinessArea";
import PL_Products from "./PL_Products";
import PL_Projects from "./PL_Projects";
import PL_ActivityPlan from "./PL_ActivityPlan";
import PL_ActivityPlanner from "./PL_ActivityPlanner";
import PL_Resources from "./PL_Resources";
import PL_Subject from "./PL_Subject";

const PL_MainComponent = (props: any) => {
  let plbaObject = {
    BA: "",
    Subject: "",
    Product: "",
    ProductVersion: "",
    ProductId: "",
    Project: "",
    ProjectVersion: "",
    ProjectId: "",
    ActivityPlan: "",
    ActivityPlanId: null,
    Page: "BusinessArea",
  };

  const [PLBAObject, setPLBAObject] = useState(plbaObject);

  const selectPLFunction = (Page, key, text, version, id) => {
    let tempObj = { ...PLBAObject };
    if (key) {
      tempObj[key] = text;
      text == "Product" || "Project"
        ? ((tempObj[key + "Version"] = version), (tempObj[key + "Id"] = id))
        : "";
      text == "ActivityPlan" ? (tempObj[key + "Id"] = id) : "";
    }
    Page ? (tempObj["Page"] = Page) : "";
    setPLBAObject(tempObj);
  };

  return (
    <>
      {PLBAObject.Page == "BusinessArea" ? (
        <PL_BusinessArea
          context={props.context}
          spcontext={props.spcontext}
          graphContent={props.graphContent}
          URL={props.URL}
          WeblistURL={props.WeblistURL}
          PLBAObject={PLBAObject}
          selectPLFunction={selectPLFunction}
        />
      ) : PLBAObject.Page == "Subject" ? (
        <PL_Subject
          context={props.context}
          spcontext={props.spcontext}
          graphContent={props.graphContent}
          URL={props.URL}
          WeblistURL={props.WeblistURL}
          PLBAObject={PLBAObject}
          selectPLFunction={selectPLFunction}
        />
      ) : PLBAObject.Page == "Product" ? (
        <PL_Products
          context={props.context}
          spcontext={props.spcontext}
          graphContent={props.graphContent}
          URL={props.URL}
          WeblistURL={props.WeblistURL}
          PLBAObject={PLBAObject}
          selectPLFunction={selectPLFunction}
        />
      ) : PLBAObject.Page == "Project" ? (
        <PL_Projects
          context={props.context}
          spcontext={props.spcontext}
          graphContent={props.graphContent}
          URL={props.URL}
          WeblistURL={props.WeblistURL}
          PLBAObject={PLBAObject}
          selectPLFunction={selectPLFunction}
        />
      ) : PLBAObject.Page == "Resources" ? (
        <PL_Resources
          context={props.context}
          spcontext={props.spcontext}
          graphContent={props.graphContent}
          URL={props.URL}
          WeblistURL={props.WeblistURL}
          PLBAObject={PLBAObject}
          selectPLFunction={selectPLFunction}
        />
      ) : (
        ""
      )}
    </>
  );
};
export default PL_MainComponent;
