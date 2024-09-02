import * as React from "react";
import { useState, useEffect } from "react";
import { Web } from "@pnp/sp/webs";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  Icon,
  Label,
  Dropdown,
  IDropdownStyles,
  SearchBox,
  ISearchBoxStyles,
  TooltipHost,
  TooltipOverflowMode,
  PrimaryButton,
  IColumn,
} from "@fluentui/react";

import Service from "../components/Services";

import "../ExternalRef/styleSheets/Styles.css";
import styles from "./InnovationHubIntranet.module.scss";
import Pagination from "office-ui-fabric-react-pagination";

import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import CustomLoader from "./CustomLoader";
import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import { arraysEqual } from "office-ui-fabric-react/lib/Utilities";
import PL_Products from "./PL_Products";

const PL_BusinessArea = (props: any) => {
  let loggeduseremail: string = props.spcontext.pageContext.user.email;
  const sharepointWeb = Web(props.URL);
  const ListNameURL = props.WeblistURL;

  const buttonStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "2px",
    fontSize: 16,
  });

  const buttonStyleClass = mergeStyleSets({
    buttonPrimary: [
      {
        color: "rgb(0 191 198)",
        backgroundColor: "rgb(227 246 246)",
        borderRadius: "3px",
        border: "none",
        marginRight: "10px",
        marginBottom: "10px",
        width: "250px",
        height: "60px",
        selectors: {
          ":hover": {
            backgroundColor: "rgb(227 246 246)",
            color: "rgb(0 191 198)",
            opacity: 0.9,
            borderRadius: "3px",
            border: "none",
            marginRight: "10px",
          },
        },
      },
      buttonStyle,
    ],
    buttonSecondary: [
      {
        color: "White",
        backgroundColor: "#038387",
        borderRadius: "3px",
        border: "none",
        margin: "0 5px",
        selectors: {
          ":hover": {
            backgroundColor: "#038387",
            opacity: 0.9,
          },
        },
      },
      buttonStyle,
    ],
  });

  // Use State
  const [PLBAReRender, setPLBAReRender] = useState(false);
  const [PLBAMaster, setPLBAMaster] = useState([]);
  const [PLBALoader, setPLBALoader] = useState("noLoader");

  const getBusinessArea = () => {
    let BAChoices = [];

    const _sortFilterKeys = (a, b) => {
      if (a.key.toLowerCase() < b.key.toLowerCase()) {
        return -1;
      }
      if (a.key.toLowerCase() > b.key.toLowerCase()) {
        return 1;
      }
      return 0;
    };

    //Business Area Choices
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .fields.getByInternalNameOrTitle("BusinessArea")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              BAChoices.findIndex((baOptn) => {
                return baOptn.key == choice;
              }) == -1
            ) {
              BAChoices.push({
                Name: choice,
                Display: choice.includes("PD")
                  ? choice.replace("PD", "")
                  : choice.includes("SS")
                  ? choice.replace("SS", "")
                  : choice.includes("SD")
                  ? choice.replace("SD", "")
                  : "",
                Type: choice.includes("PD")
                  ? "Products"
                  : choice.includes("SS") || choice.includes("SD")
                  ? "Solutions"
                  : "",
              });
            }
          }
        });
      })
      .then(() => {
        BAChoices.sort(_sortFilterKeys);
        console.log(BAChoices);
        setPLBAMaster([...BAChoices]);
        setPLBALoader("noLoader");
      })
      .catch((err) => {
        ErrorFunction(err, "getBusinessArea");
      });
  };
  const BusinessArea = () => {
    const BACollection = [
      {
        Name: "PD Curriculum",
        ShortName: "PDC",
        Display: "Curriculum",
        Type: "Products",
      },
      {
        Name: "PD Professional Learning",
        ShortName: "PDPL",
        Display: "Professional Learning",
        Type: "Products",
      },
      {
        Name: "PD School Improvements",
        ShortName: "PDSI",
        Display: "School Improvements",
        Type: "Products",
      },
      {
        Name: "SS Business",
        ShortName: "SSB",
        Display: "Business",
        Type: "Solutions",
      },
      {
        Name: "SS Publishing",
        ShortName: "SSP",
        Display: "Publishing",
        Type: "Solutions",
      },
      {
        Name: "SS Content Creation",
        ShortName: "SSCC",
        Display: "Content Creation",
        Type: "Solutions",
      },
      {
        Name: "SS Marketing",
        ShortName: "SSM",
        Display: "Marketing",
        Type: "Solutions",
      },
      {
        Name: "SS Technology",
        ShortName: "SST",
        Display: "Technology",
        Type: "Products",
      },
      {
        Name: "SS Research and Evaluation",
        ShortName: "SSRE",
        Display: "Research and Evaluation",
        Type: "Solutions",
      },
      {
        Name: "SD School Partnerships",
        ShortName: "SSPSP",
        Display: "School Partnerships",
        Type: "Solutions",
      },
    ];

    setPLBAMaster([...BACollection]);
    setPLBALoader("noLoader");
  };
  const ErrorFunction = (error: any, functionName: string) => {
    console.log(error);
    setPLBALoader("noLoader");

    // let response = {
    //   ComponentName: "PL_BusinessArea",
    //   FunctionName: functionName,
    //   ErrorMessage: JSON.stringify(error["message"]),
    //   Title: loggeduseremail,
    // };

    // Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
    //   () => {
    ErrorPopup();
    //   }
    // );
  };
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );

  //Use Effect
  useEffect(() => {
    setPLBALoader("startUpLoader");
    // getBusinessArea();

    BusinessArea();
  }, [PLBAReRender]);

  return (
    <div style={{ padding: "5px 15px" }}>
      {PLBALoader == "startUpLoader" ? (
        <CustomLoader />
      ) : (
        <>
          <div
            style={{
              display: "flex",
              alignItems: "flex-start",
              justifyContent: "space-between",
              marginBottom: 20,
              // color: "#2392b2",
            }}
          >
            <div className={styles.dpTitle}>
              <Label style={{ fontSize: 24, padding: 0 }}>Product List</Label>
            </div>
          </div>
          <div style={{ padding: "25px 75px" }}>
            <Label
              style={{
                color: "#2392b2",
                fontSize: 24,
                padding: 0,
                marginBottom: 15,
              }}
            >
              Products
            </Label>

            <div style={{ display: "flex", flexWrap: "wrap" }}>
              {PLBAMaster.filter((arr) => {
                return arr.Type == "Products";
              }).map((arr) => {
                return (
                  <div>
                    <PrimaryButton
                      text={arr.Display}
                      className={buttonStyleClass.buttonPrimary}
                      onClick={(_) => {
                        props.selectPLFunction(
                          "Subject",
                          "BA",
                          arr.Name,
                          "",
                          null
                        );
                      }}
                    ></PrimaryButton>
                  </div>
                );
              })}
            </div>
          </div>
          <div style={{ padding: "25px 75px" }}>
            <Label
              style={{
                color: "#2392b2",
                fontSize: 24,
                padding: 0,
                marginBottom: 15,
              }}
            >
              Solutions
            </Label>

            <div style={{ display: "flex", flexWrap: "wrap" }}>
              {PLBAMaster.filter((arr) => {
                return arr.Type == "Solutions";
              }).map((arr) => {
                return (
                  <div>
                    <PrimaryButton
                      text={arr.Display}
                      className={buttonStyleClass.buttonPrimary}
                      onClick={(_) => {
                        props.selectPLFunction(
                          "Subject",
                          "BA",
                          arr.Name,
                          "",
                          null
                        );
                      }}
                    ></PrimaryButton>
                  </div>
                );
              })}
            </div>
          </div>
        </>
      )}
    </div>
  );
};
export default PL_BusinessArea;
