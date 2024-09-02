import * as React from "react";
import { useState, useEffect } from "react";
import { Web } from "@pnp/sp/webs";
import { Label, ILabelStyles } from "@fluentui/react";
import "../ExternalRef/styleSheets/Styles.css";
import styles from "./InnovationHubIntranet.module.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";

import OrgAllReports from "./OrgRep";
import DocumentReview from "./DocReviewRep";
import BusinessArea from "./WeeklyProdRep";
interface IProps {
  context: WebPartContext;
  spcontext: any;
  graphContent: any;
  URL: string;
  handleclick: any;
  pageType: string;
  peopleList: any;
  isAdmin: boolean;
}

const ReportsMain = (props: IProps): JSX.Element => {
  // Style-Section Starts
  const headingStyles: Partial<ILabelStyles> = {
    root: {
      color: "#000",
      fontSize: 26,
      padding: 0,
      marginBottom: 10,
    },
  };
  // Style-Section Ends
  //State Declaration Starts
  const [ORPageSwitch, setORPageSwitch] = useState("AllReports");
  //State Declaration Ends

  useEffect(() => { }, []);
  return (
    <div style={{ padding: "10px 15px" }}>
      {/* header-Section Starts*/}
      <div>
        <div>
          <Label styles={headingStyles}>All Reports</Label>
        </div>
        <div className={styles.orgButtonSection}>
          <button
            className={
              ORPageSwitch == "AllReports"
                ? styles.activeButton
                : styles.inactiveButton
            }
            onClick={() => {
              setORPageSwitch("AllReports");
            }}
          >
            Production Report
          </button>
          <button
            className={
              ORPageSwitch == "MyReports"
                ? styles.activeButton
                : styles.inactiveButton
            }
            onClick={() => {
              setORPageSwitch("MyReports");
            }}
          >
            Organization Report
          </button>
          <button
            className={
              ORPageSwitch == "ApprovalRequests"
                ? styles.activeButton
                : styles.inactiveButton
            }
            onClick={() => {
              setORPageSwitch("ApprovalRequests");
            }}
          >
            Document Review
          </button>
        </div>
      </div>
      {/* header-Section Ends*/}
      {/* body-Section Starts */}
      <div>
        {ORPageSwitch == "AllReports" ? (
          <BusinessArea
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={props.URL}
            peopleList={props.peopleList}
            isAdmin={props.isAdmin}
            handleclick={props.handleclick}

          />
        ) : ORPageSwitch == "MyReports" ? (
          <OrgAllReports
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={props.URL}
            peopleList={props.peopleList}
            isAdmin={props.isAdmin}
          />
        ) : ORPageSwitch == "ApprovalRequests" ? (
          <DocumentReview
            context={props.context}
            spcontext={props.spcontext}
            graphContent={props.graphContent}
            URL={props.URL}
            peopleList={props.peopleList}
            isAdmin={props.isAdmin}
            handleclick={props.handleclick}
          // distributionListHandlerFunction={distributionListHandler}
          />
        ) : null}
      </div>
      {/* body-Section Ends */}
    </div>
  );
};

export default ReportsMain;
