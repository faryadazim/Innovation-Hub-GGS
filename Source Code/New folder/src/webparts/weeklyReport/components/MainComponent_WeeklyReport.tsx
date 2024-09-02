import * as React from "react";
import { useState, useEffect } from "react";
import "@pnp/sp/webs";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import {
  Label,
  ILabelStyles,
  Dropdown,
  IDropdownStyles,
} from "@fluentui/react";
import styles from "./WeeklyReport.module.scss";
import "../ExternalRef/styleSheets/Styles.css";

import WRDashboard from "./WRDashboard";
import DocumentReviewHistory from "./WRHistoryPage";
import WRDeliverable from "./WRDeliverable";
import WRDocumentreview from "./WRDocumentreview";
import WRProjectReport from "./WRProjectReport";

// /* Development URL */
// const webURL = "https://ggsaus.sharepoint.com/sites/Intranet_Test";
// const WeblistURL = "Annual Plan Test";

// /* Production URL */
// const webURL = 'https://ggsaus.sharepoint.com'
// const WeblistURL = 'Annual Plan'

var webURL;
var WeblistURL;

if (window.location.href.toLowerCase().indexOf("production/") > -1) {
  /* Production URL */
  webURL = "https://ggsaus.sharepoint.com";
  WeblistURL = "Annual Plan";
} else {
  /* Development URL */
  webURL = "https://ggsaus.sharepoint.com/sites/Intranet_Test";
  WeblistURL = "Annual Plan Test";
}

let allPeoples: any[] = [];
const MainComponent_WeeklyReport = (props: any) => {
  // Variable Declaration Starts
  const _webURL = Web(webURL);
  // Variable Declaration Starts

  // Style-Section Starts
  const headingStyles: Partial<ILabelStyles> = {
    root: {
      color: "#000",
      fontSize: 26,
      padding: 0,
      marginBottom: 10,
    },
  };
  const WRActiveDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
      outline: "none",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#badbe8",
      fontSize: 12,
      color: "#01595c",
      border: "none",
      borderBottom: "2px solid #18677e",
      borderRadius: 4,
      fontWeight: 600,
      outline: "none",
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#01595c" },
  };
  // Style-Section Ends

  // State Declaration Starts
  const [pageSwitch, setPageSwitch] = useState("Dashboard");
  const [siteUsers, setSiteUsers] = useState(allPeoples);
  const [historyData, setHistory] = useState({
    condition: false,
    sourcePage: "",
    targetID: null,
  });
  const [BusinessArea, setBusinessArea] = useState({
    Options: [],
    Filterkeys: "",
  });
  // State Declaration Starts

  // Function Declaration Starts
  const historyDataHandler = (condition: boolean, targetID: number) => {
    setHistory({
      condition: condition,
      sourcePage: pageSwitch,
      targetID: targetID,
    });
  };
  // Function Declaration Starts

  const getMasterDropdown = () => {
    let tempOptions = [];
    const sortFilterKeys = (a, b) => {
      if (a.text < b.text) {
        return -1;
      }
      if (a.text > b.text) {
        return 1;
      }
      return 0;
    };

    _webURL.lists
      .getByTitle("Master User List")
      .fields.getByInternalNameOrTitle("BusinessArea")()
      .then((response) => {
        response["Choices"].forEach((choice, index) => {
          if (choice != null) {
            if (
              tempOptions.findIndex((rpb) => {
                return rpb.key == choice;
              }) == -1
            ) {
              tempOptions.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
        tempOptions.sort(sortFilterKeys);
        BusinessArea.Options = tempOptions;
        BusinessArea.Filterkeys = BusinessArea.Options[0].key;
        setBusinessArea(BusinessArea);
        getAllUsers();
      })
      .then()
      .catch((error) => {
        console.log(error);
      });
  };

  const getAllUsers = () => {
    _webURL.siteUsers().then((_allUsers) => {
      _allUsers.forEach((user) => {
        let userName = user.Title.toLowerCase();
        // if (userName.indexOf("archive") == -1) {
        allPeoples.push({
          key: 1,
          imageUrl:
            `/_layouts/15/userphoto.aspx?size=S&accountname=` + `${user.Email}`,
          text: user.Title,
          ID: user.Id,
          secondaryText: user.Email,
          isValid: true,
        });
        // }
      });
      setSiteUsers([...allPeoples]);
    });
  };

  useEffect(() => {
    getMasterDropdown();
  }, []);
  return (
    <>
      {siteUsers.length > 0 ? (
        <div style={{ padding: "10px 15px" }}>
          {/* header-Section Starts*/}
          <div>
            <div>
              <Label styles={headingStyles}>Production report</Label>
            </div>
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
              }}
            >
              <div className={styles.WRButtonSection}>
                <button
                  className={
                    pageSwitch == "Dashboard"
                      ? styles.activeButton
                      : styles.inactiveButton
                  }
                  onClick={() => {
                    setHistory({
                      condition: false,
                      sourcePage: "",
                      targetID: null,
                    });
                    setPageSwitch("Dashboard");
                  }}
                >
                  Deliverable by week
                </button>
                <button
                  className={
                    pageSwitch == "Documentreview"
                      ? styles.activeButton
                      : styles.inactiveButton
                  }
                  onClick={() => {
                    setHistory({
                      condition: false,
                      sourcePage: "",
                      targetID: null,
                    });
                    setPageSwitch("Documentreview");
                  }}
                >
                  Document review
                </button>
                <button
                  className={
                    pageSwitch == "Deliverable"
                      ? styles.activeButton
                      : styles.inactiveButton
                  }
                  onClick={() => {
                    setHistory({
                      condition: false,
                      sourcePage: "",
                      targetID: null,
                    });
                    setPageSwitch("Deliverable");
                  }}
                >
                  Deliverable
                </button>
                <button
                  className={
                    pageSwitch == "ProjectReport"
                      ? styles.activeButton
                      : styles.inactiveButton
                  }
                  onClick={() => {
                    setHistory({
                      condition: false,
                      sourcePage: "",
                      targetID: null,
                    });
                    setPageSwitch("ProjectReport");
                  }}
                >
                  Project report
                </button>
              </div>
              <div>
                <Dropdown
                  placeholder="Select an option"
                  options={BusinessArea.Options}
                  selectedKey={BusinessArea.Filterkeys}
                  styles={WRActiveDropdownStyles}
                  onChange={(e, option: any) => {
                    BusinessArea.Filterkeys = option["key"];
                    setBusinessArea({ ...BusinessArea });
                  }}
                />
              </div>
            </div>
          </div>
          {/* header-Section Ends*/}
          {/* body-Section Starts */}
          <div>
            {historyData.condition ? (
              <DocumentReviewHistory
                context={props.context}
                spcontext={props.spcontext}
                graphContent={props.graphContent}
                URL={webURL}
                peopleList={siteUsers}
                historyDataHandler={historyDataHandler}
                historyData={historyData}
              />
            ) : pageSwitch == "Dashboard" ? (
              <WRDashboard
                context={props.context}
                spcontext={props.spcontext}
                graphContent={props.graphContent}
                URL={webURL}
                peopleList={siteUsers}
                BA={BusinessArea.Filterkeys}
              />
            ) : pageSwitch == "Deliverable" ? (
              <WRDeliverable
                context={props.context}
                spcontext={props.spcontext}
                graphContent={props.graphContent}
                URL={webURL}
                peopleList={siteUsers}
                BA={BusinessArea.Filterkeys}
                ListName={WeblistURL}
              />
            ) : pageSwitch == "Documentreview" ? (
              <WRDocumentreview
                context={props.context}
                spcontext={props.spcontext}
                graphContent={props.graphContent}
                URL={webURL}
                peopleList={siteUsers}
                historyDataHandler={historyDataHandler}
                BA={BusinessArea.Filterkeys}
              />
            ) : pageSwitch == "ProjectReport" ? (
              <WRProjectReport
                context={props.context}
                spcontext={props.spcontext}
                graphContent={props.graphContent}
                URL={webURL}
                peopleList={siteUsers}
                BA={BusinessArea.Filterkeys}
                ListName={WeblistURL}
              />
            ) : null}
          </div>
          {/* body-Section Ends */}
        </div>
      ) : null}
    </>
  );
};

export default MainComponent_WeeklyReport;
