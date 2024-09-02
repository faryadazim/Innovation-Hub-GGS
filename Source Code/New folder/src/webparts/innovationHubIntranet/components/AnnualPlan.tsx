import * as React from "react";
import { useState, useEffect } from "react";
import * as moment from "moment";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";
import {
  DetailsList,
  DetailsListLayoutMode,
  IDetailsListStyles,
  SelectionMode,
  Icon,
  Label,
  ILabelStyles,
  SearchBox,
  ISearchBoxStyles,
  Dropdown,
  IDropdownStyles,
  IDropdownOption,
  NormalPeoplePicker,
  Persona,
  PersonaPresence,
  PersonaSize,
  Modal,
  DatePicker,
  IDatePickerStyles,
  Checkbox,
  ICheckboxStyles,
  TextField,
  ITextFieldStyles,
  Spinner,
  TooltipHost,
  TooltipDelay,
  TooltipOverflowMode,
  DirectionalHint,
  IColumn,
} from "@fluentui/react";

import Service from "../components/Services";

import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
import Pagination from "office-ui-fabric-react-pagination";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";
import "../ExternalRef/styleSheets/Styles.css";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import styles from "./InnovationHubIntranet.module.scss";
import CustomLoader from "./CustomLoader";

let columnSortArr = [];
let columnSortMasterArr = [];
let editYear = [];
let DateListFormat = "DD/MM/YYYY";
let DatePickerFormat = "YYYY-MM-DDT14:00:00Z";

const AnnualPlan = (props: any) => {
  const sharepointWeb = Web(props.URL);
  const ListNameURL = props.WeblistURL;

  let loggeduseremail: string = props.spcontext.pageContext.user.email;

  let currentpage = 1;
  let totalPageItems = 10;
  const apAllitems = [];
  const apMasterProductCollection = [];
  const allPeoples = props.peopleList;
  const BAacronymsCollection = [
    {
      Name: "PD Curriculum",
      ShortName: "PDC",
    },
    {
      Name: "PD Professional Learning",
      ShortName: "PDPL",
    },
    {
      Name: "PD School Improvements",
      ShortName: "PDSI",
    },
    {
      Name: "SS Business",
      ShortName: "SSB",
    },
    {
      Name: "SS Publishing",
      ShortName: "SSP",
    },
    {
      Name: "SS Content Creation",
      ShortName: "SSCC",
    },
    {
      Name: "SS Marketing",
      ShortName: "SSM",
    },
    {
      Name: "SS Technology",
      ShortName: "SST",
    },
    {
      Name: "SS Research and Evaluation",
      ShortName: "SSRE",
    },
    {
      Name: "SD School Partnerships",
      ShortName: "SSPSP",
    },
  ];
  // const TODacronymsCollection = [
  //   {
  //     Name: "Product",
  //     ShortName: "PT",
  //   },
  //   {
  //     Name: "Project",
  //     ShortName: "PR",
  //   },
  //   {
  //     Name: "Task",
  //     ShortName: "T",
  //   },
  //   {
  //     Name: "Activity",
  //     ShortName: "A",
  //   },
  //   {
  //     Name: "Product initiative",
  //     ShortName: "PTI",
  //   },
  //   {
  //     Name: "Product tool",
  //     ShortName: "PRT",
  //   },
  //   {
  //     Name: "Product related",
  //     ShortName: "TPR",
  //   },
  //   {
  //     Name: "Activity planner",
  //     ShortName: "AP",
  //   },
  //   {
  //     Name: "Organisation solution",
  //     ShortName: "SPT",
  //   },
  //   {
  //     Name: "Project solution",
  //     ShortName: "SPR",
  //   },
  //   {
  //     Name: "Task solution",
  //     ShortName: "ST",
  //   },
  //   {
  //     Name: "Activity solution",
  //     ShortName: "BA",
  //   },
  //   {
  //     Name: "Test",
  //     ShortName: "TTF",
  //   },
  //   {
  //     Name: "Product initiative",
  //     ShortName: "NI",
  //   },
  //   {
  //     Name: "Product tool",
  //     ShortName: "TEC",
  //   },
  //   {
  //     Name: "Product related",
  //     ShortName: "S",
  //   },
  //   {
  //     Name: "Organisation solution",
  //     ShortName: "OS",
  //   },
  //   {
  //     Name: "Project solution",
  //     ShortName: "SS",
  //   },
  //   {
  //     Name: "Task solution",
  //     ShortName: "TS",
  //   },
  //   {
  //     Name: "Activity solution",
  //     ShortName: "AS",
  //   },
  // ];

  const TODacronymsCollection = [
    {
      Name: "Product",
      ShortName: "PT",
    },
    {
      Name: "Project",
      ShortName: "PR",
    },
    {
      Name: "Task",
      ShortName: "T",
    },
    {
      Name: "Activity",
      ShortName: "A",
    },
    {
      Name: "Product initiative",
      ShortName: "NI",
    },
    {
      Name: "Technology",
      ShortName: "TEC",
    },
    {
      Name: "Strategy",
      ShortName: "S",
    },
    {
      Name: "Activity planner",
      ShortName: "AP",
    },
    {
      Name: "Organisation solution",
      ShortName: "OS",
    },
    {
      Name: "System solution",
      ShortName: "SS",
    },
    {
      Name: "Task solution",
      ShortName: "TS",
    },
    {
      Name: "Activity solution",
      ShortName: "AS",
    },
    {
      Name: "Test",
      ShortName: "TTF",
    },
  ];

  const apColumns = props.isAdmin
    ? [
        {
          key: "BAacronyms",
          name: "BA",
          fieldName: "BAacronyms",
          minWidth: 35,
          maxWidth: 50,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
        },
        {
          key: "Term",
          name: "Term",
          fieldName: "Term",
          minWidth: 50,
          maxWidth: 60,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => <>{item.Term.join(",")}</>,
        },
        {
          key: "Hours",
          name: "Hours",
          fieldName: "Hours",
          minWidth: 55,
          maxWidth: 60,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
        },
        {
          key: "StartDate",
          name: "Start date",
          fieldName: "StartDate",
          minWidth: 80,
          maxWidth: 100,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => item.StartDate,
        },
        {
          key: "EndDate",
          name: "End date",
          fieldName: "EndDate",
          minWidth: 75,
          maxWidth: 100,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => item.EndDate,
        },
        {
          key: "Product",
          name: "Product or solution",
          fieldName: "Product",
          minWidth: 150,
          maxWidth: 300,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <>
              <TooltipHost
                id={item.ID}
                content={item.Product}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.Product}</span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "TypeOfProject",
          name: "TOD",
          fieldName: "TypeOfProject",
          minWidth: 45,
          maxWidth: 60,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
        },
        {
          key: "ProjectOrTask",
          name: "Name of the deliverable",
          fieldName: "ProjectOrTask",
          minWidth: 220,
          maxWidth: 230,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <>
              <TooltipHost
                id={item.ID}
                content={item.ProjectOrTask + " " + item.ProjectVersion}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>
                  {item.ProjectOrTask + " " + item.ProjectVersion}
                </span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "Priority",
          name: "Priority",
          fieldName: "Priority",
          minWidth: 65,
          maxWidth: 65,
          onColumnClick: (ev, column) => {
            _onColumnClick(ev, column);
          },
        },
        {
          key: "AP",
          name: "AP",
          fieldName: "AP",
          minWidth: 30,
          maxWidth: 70,

          onRender: (item) => (
            <>
              <Icon
                style={{
                  marginLeft: 0,
                }}
                iconName="Link12"
                className={apIconStyleClass.link}
                onClick={() => {
                  props.handleclick("ActivityPlan", item.ID);
                }}
              />
            </>
          ),
        },
        {
          key: "DP/AP",
          name: "DP",
          fieldName: "DPAP",
          minWidth: 30,
          maxWidth: 70,

          onRender: (item) => (
            <>
              <Icon
                style={{
                  marginLeft: 0,
                }}
                iconName="Link12"
                className={apIconStyleClass.link}
                onClick={() => {
                  props.handleclick("DeliveryPlan", item.ID, "AP");
                }}
              />
            </>
          ),
        },
        {
          key: "PB",
          name: "PB",
          fieldName: "PB",
          minWidth: 30,
          maxWidth: 70,

          onRender: (item) => (
            <>
              <Icon
                iconName="Link12"
                className={apIconStyleClass.link}
                onClick={() => {
                  props.handleclick("ProductionBoard", item.ID, "AP");
                }}
              />
            </>
          ),
        },
        {
          key: "Status",
          name: "Status",
          fieldName: "Status",
          minWidth: 120,
          maxWidth: 120,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <>
              {item.Status == "Completed" ? (
                <div className={apStatusStyleClass.completed}>
                  {item.Status}
                </div>
              ) : item.Status == "Scheduled" ? (
                <div className={apStatusStyleClass.scheduled}>
                  {item.Status}
                </div>
              ) : item.Status == "On schedule" ? (
                <div className={apStatusStyleClass.onSchedule}>
                  {item.Status}
                </div>
              ) : item.Status == "Behind schedule" ? (
                <div className={apStatusStyleClass.behindScheduled}>
                  {item.Status}
                </div>
              ) : item.Status == "On hold" ? (
                <div className={apStatusStyleClass.Onhold}>{item.Status}</div>
              ) : (
                ""
              )}
            </>
          ),
        },
        {
          key: "PM",
          name: "C",
          fieldName: "PM",
          minWidth: 50,
          maxWidth: 80,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <>
              {item.PMName.id ? (
                <>
                  {
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "flex-start",
                        cursor: "pointer",
                      }}
                    >
                      <TooltipHost
                        content={
                          <ul style={{ margin: 10, padding: 0 }}>
                            <li>
                              <div style={{ display: "flex" }}>
                                <Persona
                                  showOverflowTooltip
                                  size={PersonaSize.size24}
                                  presence={PersonaPresence.none}
                                  showInitialsUntilImageLoads={true}
                                  imageUrl={
                                    "/_layouts/15/userphoto.aspx?size=S&username=" +
                                    `${item.PMName.email}`
                                  }
                                />
                                <div style={{ marginLeft: 10 }}>
                                  {item.PMName.name}
                                </div>
                              </div>
                            </li>
                          </ul>
                        }
                        delay={TooltipDelay.zero}
                        id={item.ID}
                        directionalHint={DirectionalHint.bottomCenter}
                        styles={{ root: { display: "inline-block" } }}
                      >
                        <Persona
                          aria-describedby={item.ID}
                          size={PersonaSize.size24}
                          presence={PersonaPresence.none}
                          imageUrl={
                            "/_layouts/15/userphoto.aspx?size=S&username=" +
                            `${item.PMName.email}`
                          }
                        />
                      </TooltipHost>
                    </div>
                  }
                </>
              ) : (
                ""
              )}
            </>
          ),
        },
        {
          key: "D",
          name: "D",
          fieldName: "D",
          minWidth: 50,
          maxWidth: 80,
          onRender: (item) => (
            <>
              {item.DNames.length > 0 ? (
                <>
                  {
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "flex-start",
                        cursor: "pointer",
                      }}
                    >
                      <div title={item.DNames[0].name}>
                        <Persona
                          showOverflowTooltip
                          size={PersonaSize.size24}
                          presence={PersonaPresence.none}
                          showInitialsUntilImageLoads={true}
                          imageUrl={
                            "/_layouts/15/userphoto.aspx?size=S&username=" +
                            `${item.DNames[0].email}`
                          }
                        />
                      </div>
                      {item.DNames.length > 1 ? (
                        <TooltipHost
                          content={
                            <ul style={{ margin: 10, padding: 0 }}>
                              {item.DNames.map((DName) => {
                                return (
                                  <li>
                                    <div style={{ display: "flex" }}>
                                      <Persona
                                        showOverflowTooltip
                                        size={PersonaSize.size24}
                                        presence={PersonaPresence.none}
                                        showInitialsUntilImageLoads={true}
                                        imageUrl={
                                          "/_layouts/15/userphoto.aspx?size=S&username=" +
                                          `${DName.email}`
                                        }
                                      />
                                      <Label style={{ marginLeft: 10 }}>
                                        {DName.name}
                                      </Label>
                                    </div>
                                  </li>
                                );
                              })}
                            </ul>
                          }
                          delay={TooltipDelay.zero}
                          id={item.ID}
                          directionalHint={DirectionalHint.bottomCenter}
                          styles={{ root: { display: "inline-block" } }}
                        >
                          <div
                            className={styles.extraPeople}
                            aria-describedby={item.ID}
                          >
                            {item.DNames.length}
                          </div>
                        </TooltipHost>
                      ) : null}
                    </div>
                  }
                </>
              ) : (
                ""
              )}
            </>
          ),
        },
        {
          key: "Actions",
          name: "Actions",
          fieldName: "Actions",
          minWidth: 65,
          maxWidth: 65,

          onRender: (item) => (
            <>
              <Icon
                title="Edit deliverable"
                iconName="Edit"
                className={apIconStyleClass.edit}
                onClick={() => {
                  editYear = [];
                  for (
                    let year = moment().year();
                    year <= moment().year() + 10;
                    year++
                  ) {
                    editYear.push({
                      key: year,
                      text: year,
                    });
                  }
                  if (
                    editYear.findIndex((yr) => {
                      return yr.key == item.Year;
                    }) == -1 &&
                    item.Year != ""
                  ) {
                    editYear.unshift({
                      key: item.Year,
                      text: item.Year,
                    });
                  }

                  let filteredArr = columnSortArr.filter((data) => {
                    return data.ID == item.ID;
                  });

                  let devs = [];
                  if (filteredArr[0].DNames.length > 0) {
                    filteredArr[0].DNames.forEach((dev) => {
                      devs.push(dev.userDetails);
                    });
                  }
                  setApResponseData({
                    ID: filteredArr[0].ID,
                    businessArea: filteredArr[0].BusinessArea,
                    typeOfProject: filteredArr[0].TypeOfProject,
                    term:
                      filteredArr[0].Term.length > 0 ? filteredArr[0].Term : [],
                    product: filteredArr[0].Product,
                    startDate: filteredArr[0].StartDate
                      ? new Date(
                          moment(
                            filteredArr[0].DefaultStartDate,
                            DateListFormat
                          ).format(DatePickerFormat)
                        )
                      : null,
                    endDate: filteredArr[0].EndDate
                      ? new Date(
                          moment(
                            filteredArr[0].DefaultEndDate,
                            DateListFormat
                          ).format(DatePickerFormat)
                        )
                      : null,
                    projectOrTask: filteredArr[0].ProjectOrTask,
                    ProjectVersion: filteredArr[0].ProjectVersion,
                    Priority: filteredArr[0].Priority,
                    year: filteredArr[0].Year,
                    manager: filteredArr[0].PMName.id,
                    developer: devs,
                    status: filteredArr[0].Status,
                  });
                  setApModalBoxVisibility({
                    condition: true,
                    action: "Update",
                    selectedItem: filteredArr,
                  });
                }}
              />
              <Icon
                iconName="Delete"
                title="Delete deliverable"
                className={apIconStyleClass.delete}
                onClick={() => {
                  setApDeletePopup({ condition: true, targetId: item.ID });
                }}
              />
            </>
          ),
        },
      ]
    : [
        {
          key: "BAacronyms",
          name: "BA",
          fieldName: "BAacronyms",
          minWidth: 30,
          maxWidth: 50,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
        },
        {
          key: "Term",
          name: "T",
          fieldName: "Term",
          minWidth: 50,
          maxWidth: 60,
          onRender: (item) => <>{item.Term.join(",")}</>,
        },
        {
          key: "Hours",
          name: "Hours",
          fieldName: "Hours",
          minWidth: 55,
          maxWidth: 60,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
        },
        {
          key: "StartDate",
          name: "Start date",
          fieldName: "StartDate",
          minWidth: 70,
          maxWidth: 100,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => item.StartDate,
        },
        {
          key: "EndDate",
          name: "End date",
          fieldName: "EndDate",
          minWidth: 70,
          maxWidth: 100,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => item.EndDate,
        },
        {
          key: "Product",
          name: "Product or solution",
          fieldName: "Product",
          minWidth: 200,
          maxWidth: 350,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <>
              <TooltipHost
                id={item.ID}
                content={item.Product}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.Product}</span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "TypeOfProject",
          name: "TOD",
          fieldName: "TypeOfProject",
          minWidth: 45,
          maxWidth: 70,
        },
        {
          key: "ProjectOrTask",
          name: "Name of the deliverable",
          fieldName: "ProjectOrTask",
          minWidth: 200,
          maxWidth: 350,
          onRender: (item) => (
            <>
              <TooltipHost
                id={item.ID}
                content={item.ProjectOrTask}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.ProjectOrTask}</span>
              </TooltipHost>
            </>
          ),
        },
        {
          key: "Priority",
          name: "Priority",
          fieldName: "Priority",
          minWidth: 65,
          maxWidth: 65,
          onColumnClick: (ev, column) => {
            _onColumnClick(ev, column);
          },
        },
        {
          key: "DP/AP",
          name: "DP",
          fieldName: "DPAP",
          minWidth: 30,
          maxWidth: 70,

          onRender: (item) => (
            <>
              <Icon
                style={{
                  marginLeft: 0,
                }}
                iconName="Link12"
                className={apIconStyleClass.link}
                onClick={() => {
                  props.handleclick("DeliveryPlan", item.ID, "AP");
                }}
              />
            </>
          ),
        },
        {
          key: "PB",
          name: "PB",
          fieldName: "PB",
          minWidth: 30,
          maxWidth: 70,

          onRender: (item) => (
            <>
              <Icon
                iconName="Link12"
                className={apIconStyleClass.link}
                onClick={() => {
                  props.handleclick("ProductionBoard", item.ID, "AP");
                }}
              />
            </>
          ),
        },
        {
          key: "Status",
          name: "Status",
          fieldName: "Status",
          minWidth: 120,
          maxWidth: 120,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <>
              {item.Status == "Completed" ? (
                <div className={apStatusStyleClass.completed}>
                  {item.Status}
                </div>
              ) : item.Status == "Scheduled" ? (
                <div className={apStatusStyleClass.scheduled}>
                  {item.Status}
                </div>
              ) : item.Status == "On schedule" ? (
                <div className={apStatusStyleClass.onSchedule}>
                  {item.Status}
                </div>
              ) : item.Status == "Behind schedule" ? (
                <div className={apStatusStyleClass.behindScheduled}>
                  {item.Status}
                </div>
              ) : item.Status == "On hold" ? (
                <div className={apStatusStyleClass.Onhold}>{item.Status}</div>
              ) : (
                ""
              )}
            </>
          ),
        },
        {
          key: "PM",
          name: "Client",
          fieldName: "PM",
          minWidth: 50,
          maxWidth: 80,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <>
              {item.PMName.id ? (
                <>
                  {
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "flex-start",
                        cursor: "pointer",
                      }}
                    >
                      <TooltipHost
                        content={
                          <ul style={{ margin: 10, padding: 0 }}>
                            <li>
                              <div style={{ display: "flex" }}>
                                <Persona
                                  showOverflowTooltip
                                  size={PersonaSize.size24}
                                  presence={PersonaPresence.none}
                                  showInitialsUntilImageLoads={true}
                                  imageUrl={
                                    "/_layouts/15/userphoto.aspx?size=S&username=" +
                                    `${item.PMName.email}`
                                  }
                                />
                                <div style={{ marginLeft: 10 }}>
                                  {item.PMName.name}
                                </div>
                              </div>
                            </li>
                          </ul>
                        }
                        delay={TooltipDelay.zero}
                        id={item.ID}
                        directionalHint={DirectionalHint.bottomCenter}
                        styles={{ root: { display: "inline-block" } }}
                      >
                        <Persona
                          aria-describedby={item.ID}
                          size={PersonaSize.size24}
                          presence={PersonaPresence.none}
                          imageUrl={
                            "/_layouts/15/userphoto.aspx?size=S&username=" +
                            `${item.PMName.email}`
                          }
                        />
                      </TooltipHost>
                    </div>
                  }
                </>
              ) : (
                ""
              )}
            </>
          ),
        },
        {
          key: "D",
          name: "Developer",
          fieldName: "D",
          minWidth: 50,
          maxWidth: 80,
          onRender: (item) => (
            <>
              {item.DNames.length > 0 ? (
                <>
                  {
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "flex-start",
                        cursor: "pointer",
                      }}
                    >
                      <div title={item.DNames[0].name}>
                        <Persona
                          showOverflowTooltip
                          size={PersonaSize.size24}
                          presence={PersonaPresence.none}
                          showInitialsUntilImageLoads={true}
                          imageUrl={
                            "/_layouts/15/userphoto.aspx?size=S&username=" +
                            `${item.DNames[0].email}`
                          }
                        />
                      </div>
                      {item.DNames.length > 1 ? (
                        <TooltipHost
                          content={
                            <ul style={{ margin: 10, padding: 0 }}>
                              {item.DNames.map((DName) => {
                                return (
                                  <li>
                                    <div style={{ display: "flex" }}>
                                      <Persona
                                        showOverflowTooltip
                                        size={PersonaSize.size24}
                                        presence={PersonaPresence.none}
                                        showInitialsUntilImageLoads={true}
                                        imageUrl={
                                          "/_layouts/15/userphoto.aspx?size=S&username=" +
                                          `${DName.email}`
                                        }
                                      />
                                      <Label style={{ marginLeft: 10 }}>
                                        {DName.name}
                                      </Label>
                                    </div>
                                  </li>
                                );
                              })}
                            </ul>
                          }
                          delay={TooltipDelay.zero}
                          id={item.ID}
                          directionalHint={DirectionalHint.bottomCenter}
                          styles={{ root: { display: "inline-block" } }}
                        >
                          <div
                            className={styles.extraPeople}
                            aria-describedby={item.ID}
                          >
                            {item.DNames.length}
                          </div>
                        </TooltipHost>
                      ) : null}
                    </div>
                  }
                </>
              ) : (
                ""
              )}
            </>
          ),
        },
      ];
  const apDrpDwnOptns = {
    baOptns: [{ key: "All", text: "All" }],
    todOptns: [{ key: "All", text: "All" }],
    potOptns: [{ key: "All", text: "All" }],
    managerOptns: [{ key: "All", text: "All" }],
    PriorityOptns: [{ key: "All", text: "All" }],
    developerOptns: [{ key: "All", text: "All" }],
    termOptns: [{ key: "All", text: "All" }],
    yearOptns: [{ key: "All", text: "All" }],
  };
  const apModalBoxDrpDwnOptns = {
    baOptns: [],
    todOptns: [],
    potOptns: [],
    managerOptns: [],
    developerOptns: [],
    PriorityOptns: [],
    termOptns: [],
    productOptns: [],
    yearOptns: [],
    statusOtpns: [
      { key: "On hold", text: "On hold" },
      { key: "Completed", text: "Completed" },
    ],
  };
  const apFilterKeys = {
    ProjectOrTaskSearch: "",
    BusinessArea: "All",
    TypeOfProject: "All",
    ProjectOrTask: "All",
    PM: "All",
    D: "All",
    Term: "All",
    Year: "All",
  };
  const responseData = {
    ID: null,
    businessArea: "",
    typeOfProject: "",
    term: [],
    product: "",
    startDate: new Date(),
    endDate: new Date(),
    projectOrTask: "",
    Priority: "",
    ProjectVersion: "",
    year: "",
    manager: "",
    developer: [],
    status: "",
  };
  const apErrorStatus = {
    businessAreaError: "",
    projectOrTaskError: "",
    productError: "",
  };

  //StylesStart

  const gridStyles: Partial<IDetailsListStyles> = {
    root: {
      selectors: {
        ".ms-DetailsRow-fields": { minHeight: 38 },
        "& [role=grid]": {
          display: "flex",
          flexDirection: "column",
          alignItems: "start",
          ".ms-DetailsRow-cell": {
            height: 38,
            minHeight: 38,
            padding: "11px 12px",
          },
        },
      },
    },
    headerWrapper: {
      flex: "0 0 auto",
    },
    contentWrapper: {
      flex: "1 1 auto",
      overflowY: "auto",
      overflowX: "hidden",
    },
  };
  const apLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const apShortLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 75,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const apSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
    },
    icon: { fontSize: 14, color: "#000" },
  };
  const apActiveSearchBoxStyles: Partial<ISearchBoxStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
      outline: "none",
      color: "#038387",
      border: "2px solid #038387",
      borderRadius: "4px",
    },
    icon: { fontSize: 14, color: "#038387" },
  };
  const apDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#7C7C7C",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    callout: {
      maxHeight: "400px !important",
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const apActiveDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 186,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#038387",
      border: "2px solid #038387",
      borderRadius: "4px",
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    callout: {
      maxHeight: "400px !important",
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const apShortDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 75,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#7C7C7C",
      border: "1px solid #E8E8EA",
      borderRadius: "4px",
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000" },
  };
  const apActiveShortDropdownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 75,
      marginRight: "15px",
      backgroundColor: "#F5F5F7",
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#038387",
      border: "2px solid #038387",
      borderRadius: 4,
      fontWeight: 600,
    },
    dropdownItem: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    caretDown: { fontSize: 14, color: "#000" },
    callout: {
      maxHeight: 300,
    },
  };
  const apModalBoxDropdownStyles: Partial<IDropdownStyles> = {
    root: { width: "300px", margin: "10px 20px" },
    title: {
      fontSize: 12,
      borderRadius: "4px",
      border: "1px solid",
      height: "36px",
      padding: "3px 10px",
      color: "#000",
    },
    dropdownItemsWrapper: { fontSize: 12 },
    dropdownItemSelected: { fontSize: 12 },
    caretDown: {
      fontSize: 14,
      padding: "3px",
      color: "#000",
      fontWeight: "bold",
    },
  };
  const apModalBoxDrpDwnCalloutStyles: Partial<IDropdownStyles> = {
    root: { width: "300px", margin: "10px 20px" },
    title: {
      fontSize: 12,
      borderRadius: "4px",
      border: "1px solid",
      padding: "3px 10px",
      height: "36px",
      color: "#000",
    },
    dropdownItemsWrapper: { fontSize: 12 },
    dropdownItemSelected: { fontSize: 12 },
    caretDown: {
      fontSize: 14,
      paddingTop: "3px",
      color: "#000",
      fontWeight: "bold",
    },
    callout: { height: 200 },
  };
  const apTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: { width: "780px", margin: "10px 20px", borderRadius: "4px" },
    field: { fontSize: 12, color: "#000" },
  };
  const apTxtBoxStylesSmall: Partial<ITextFieldStyles> = {
    root: { width: "100px", margin: "10px 10px", borderRadius: "4px" },
    field: { fontSize: 12, color: "#000" },
  };
  const apModalBoxDatePickerStyles: Partial<IDatePickerStyles> = {
    root: {
      width: "300px",
      margin: "10px 20px",
      borderRadius: "4px",
    },
    icon: {
      fontSize: "17px",
      color: "#000",
      fontWeight: "bold",
    },
  };
  const apModalBoxCheckBoxStyles: Partial<ICheckboxStyles> = {
    root: { marginTop: "46px", transform: "translateX(-26px)" },
    label: { fontWeight: "600" },
  };
  const apModalBoxPP = mergeStyles({
    width: "300px",
    margin: "10px 20px",
  });
  const apIconStyle = mergeStyles({
    fontSize: 17,
    height: 14,
    width: 17,
    cursor: "pointer",
  });
  const apIconStyleClass = mergeStyleSets({
    link: [{ color: "#2392B2", margin: "0" }, apIconStyle],
    delete: [{ color: "#CB1E06", margin: "0 7px " }, apIconStyle],
    edit: [{ color: "#2392B2", margin: "0 7px 0 0" }, apIconStyle],
    refresh: [
      {
        color: "white",
        fontSize: "18px",
        height: 20,
        width: 22,
        cursor: "pointer",
        backgroundColor: "#038387",
        padding: 5,
        marginTop: 27,
        borderRadius: 2,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        ":hover": {
          backgroundColor: "#025d60",
        },
      },
    ],
    export: [
      {
        color: "black",
        fontSize: "18px",
        height: 20,
        width: 20,
        cursor: "pointer",
        marginRight: 5,
      },
    ],
  });
  const apStatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
  });
  const apStatusStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#438700",
        backgroundColor: "#D9FFB3",
      },
      apStatusStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#9f6700",
        backgroundColor: "#FFDB99",
      },
      apStatusStyle,
    ],
    onSchedule: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#B3B300 ",
        backgroundColor: "#FFFFB3",
      },
      apStatusStyle,
    ],
    behindScheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#FF0000",
        backgroundColor: "#FFB3B3",
      },
      apStatusStyle,
    ],
    Onhold: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#773030",
        backgroundColor: "#e6b1b1",
      },
      apStatusStyle,
    ],
  });

  //stylesEnd

  const [apReRender, setapReRender] = useState(true);
  const [apUnsortMasterData, setApUnsortMasterData] = useState(apAllitems);
  const [apMasterData, setApMasterData] = useState(apAllitems);
  const [apData, setApData] = useState(apAllitems);
  const [displayData, setdisplayData] = useState(apAllitems);
  const [apResponseData, setApResponseData] = useState(responseData);
  const [apMasterProducts, setApMasterProducts] = useState(
    apMasterProductCollection
  );
  const [apDropDownOptions, setApDropDownOptions] = useState(apDrpDwnOptns);
  const [apModalBoxDropDownOptions, setApModalBoxDropDownOptions] = useState(
    apModalBoxDrpDwnOptns
  );
  const [apFilterOptions, setApFilterOptions] = useState(apFilterKeys);
  const [apModalBoxVisibility, setApModalBoxVisibility] = useState({
    condition: false,
    action: "",
    selectedItem: [],
  });
  const [apDeletePopup, setApDeletePopup] = useState({
    condition: false,
    targetId: 0,
  });
  const [submitConfirmationPopup, setSubmitConfirmationPopup] = useState(false);
  const [apModelBoxDrpDwnToTxtBox, setApModelBoxDrpDwnToTxtBox] =
    useState(false);
  const [apcurrentPage, setApCurrentPage] = useState(currentpage);
  const [apShowMessage, setApShowMessage] = useState(apErrorStatus);
  const [apStartUpLoader, setApStartUpLoader] = useState(true);
  const [apOnSubmitLoader, setApOnSubmitLoader] = useState(false);
  const [apOnDeleteLoader, setApOnDeleteLoader] = useState(false);
  const [apSubmitConfirmLoader, setApSubmitConfirmLoader] = useState(false);
  const [masterApColumn, setMasterApColumn] = useState(apColumns);

  const getApData = () => {
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.select(
        "*",
        "ProjectOwner/Title",
        "ProjectOwner/Id",
        "ProjectOwner/EMail",
        "ProjectLead/Title",
        "ProjectLead/Id",
        "ProjectLead/EMail",
        "Master_x0020_Project/Title",
        "Master_x0020_Project/Id",
        "Master_x0020_Project/ProductVersion",
        "FieldValuesAsText/StartDate",
        "FieldValuesAsText/PlannedEndDate"
      )
      .expand(
        "ProjectOwner",
        "ProjectLead",
        "Master_x0020_Project",
        "FieldValuesAsText"
      )
      .top(5000)
      .orderBy("Modified", false)
      .get()
      .then((items) => {
        items.forEach((item: any, index: number) => {
          allitemsArrayFormatter(item, apAllitems);
        });
        filterKeys(items);
        setApUnsortMasterData(apAllitems);
        columnSortArr = apAllitems;
        setApData(apAllitems);
        columnSortMasterArr = apAllitems;
        setApMasterData(apAllitems);
        setApStartUpLoader(false);
        paginate(1);
        const pageName = new URLSearchParams(window.location.search).get("TOD");
        pageName ? checkForTOD(pageName) : null;
      })
      .catch((err) => {
        apErrorFunction(err, "getApData");
      });
  };
  const checkForTOD = (todType) => {
    setApResponseData({
      ID: null,
      businessArea: "",
      typeOfProject:
        apModalBoxDropDownOptions.todOptns.filter((option) => {
          return option.key == todType;
        }).length > 0
          ? apModalBoxDropDownOptions.todOptns.filter((option) => {
              return option.key == todType;
            })[0].key
          : "",
      term: [],
      product: "",
      startDate: new Date(),
      endDate: new Date(),
      projectOrTask: "",
      Priority: "",
      ProjectVersion: "",
      year: "",
      manager: "",
      developer: [],
      status: "",
    });
    setApModalBoxVisibility({
      condition: true,
      action: "Add",
      selectedItem: [],
    });
  };
  const allitemsArrayFormatter = (item, allItems) => {
    let apDevelopersNames = [];
    let arrTerm = [];
    arrTerm.push(`${item.Term}`);
    if (item.ProjectLeadId != null) {
      item.ProjectLead.forEach((dev) => {
        apDevelopersNames.push({
          name: dev.Title,
          id: dev.Id,
          email: dev.EMail,
          userDetails: allPeoples.filter((people) => {
            return people.ID == dev.Id;
          })[0],
        });
      });
    } else {
      apDevelopersNames.push({
        name: null,
        id: null,
        email: null,
      });
    }

    allItems.push({
      ID: item.ID ? item.ID : "",
      Hours: item.AllocatedHours ? item.AllocatedHours : "",
      DefaultStartDate: item.StartDate
        ? moment(item["FieldValuesAsText"].StartDate, DateListFormat).format(
            DateListFormat
          )
        : "",
      StartDate: item.StartDate
        ? moment(item["FieldValuesAsText"].StartDate, DateListFormat).format(
            DateListFormat
          )
        : "",
      DefaultEndDate: item.PlannedEndDate
        ? moment(
            item["FieldValuesAsText"].PlannedEndDate,
            DateListFormat
          ).format(DateListFormat)
        : "",
      EndDate: item.PlannedEndDate
        ? moment(
            item["FieldValuesAsText"].PlannedEndDate,
            DateListFormat
          ).format(DateListFormat)
        : "",
      Product: item.Master_x0020_ProjectId
        ? item.Master_x0020_Project.Title +
          " " +
          (item.Master_x0020_Project.ProductVersion
            ? item.Master_x0020_Project.ProductVersion
            : "V1")
        : "",
      TypeOfProject: item.ProjectType ? item.ProjectType : "",
      Year: item.Year ? item.Year : "",
      // Term:
      //   item.TermNew != null && item.TermNew.length > 0
      //     ? [...item.TermNew]
      //     : [],
      Term:
        item.TermNew != null && item.TermNew.length > 0
          ? [...item.TermNew]
          : item.Term
          ? [...arrTerm]
          : [],
      BusinessArea: item.BusinessArea ? item.BusinessArea : "",
      BAacronyms: item.BA_x0020_acronyms ? item.BA_x0020_acronyms : "",
      ProjectOrTask: item.Title ? item.Title : "",
      Status: item.Status ? item.Status : "",
      StatusStage: item.Status,
      DPAP: "",
      PMName:
        item.ProjectOwnerId != null
          ? {
              name: item.ProjectOwner.Title,
              id: item.ProjectOwner.Id,
              email: item.ProjectOwner.EMail,
            }
          : {
              name: null,
              id: null,
              email: null,
            },
      Priority: item.Priority ? item.Priority : "",
      ProjectVersion: item.ProjectVersion ? item.ProjectVersion : "V1",
      DNames: item.ProjectLeadId != null ? [...apDevelopersNames] : [],
    });

    return allItems;
  };
  const getAllOptions = () => {
    const _sortFilterKeys = (a, b) => {
      if (a.text.toLowerCase() < b.text.toLowerCase()) {
        return -1;
      }
      if (a.text.toLowerCase() > b.text.toLowerCase()) {
        return 1;
      }
      return 0;
    };

    //Product Choices
    sharepointWeb.lists
      .getByTitle("Master Product List")
      .items.filter("IsDeleted ne 1")
      .top(5000)
      .get()
      .then((allProducts) => {
        allProducts.forEach((product) => {
          if (product.Title != null) {
            if (
              apModalBoxDrpDwnOptns.productOptns.findIndex((productOptn) => {
                return productOptn.text == product.Title;
              }) == -1
            ) {
              if (product.Title != "Not Sure") {
                apModalBoxDrpDwnOptns.productOptns.push({
                  key: product.Title + " " + product.ProductVersion,
                  text: product.Title + " " + product.ProductVersion,
                });
              }
              apMasterProductCollection.push({
                productName: product.Title,
                ProductId: product.Id,
                ProductKey: product.Title + " " + product.ProductVersion,
              });
            }
          }
        });
      })
      .then(() => {
        apModalBoxDrpDwnOptns.productOptns.sort(_sortFilterKeys);
        apModalBoxDrpDwnOptns.productOptns.unshift({
          key: "Not Sure V1",
          text: "Not Sure V1",
        });
      })
      .catch((err) => {
        apErrorFunction(err, "getAllOptions-Product");
      });

    //Business Area Choices
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .fields.getByInternalNameOrTitle("BusinessArea")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              apModalBoxDrpDwnOptns.baOptns.findIndex((baOptn) => {
                return baOptn.key == choice;
              }) == -1
            ) {
              apModalBoxDrpDwnOptns.baOptns.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then(() => {
        apModalBoxDrpDwnOptns.baOptns.sort(_sortFilterKeys);
      })
      .catch((err) => {
        apErrorFunction(err, "getAllOptions-Business Area");
      });

    //Priority  Choices
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .fields.getByInternalNameOrTitle("Priority")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              apModalBoxDrpDwnOptns.PriorityOptns.findIndex((PriorityOptn) => {
                return PriorityOptn.key == choice;
              }) == -1
            ) {
              apModalBoxDrpDwnOptns.PriorityOptns.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then(() => {
        apModalBoxDrpDwnOptns.PriorityOptns.sort(_sortFilterKeys);
      })
      .catch((err) => {
        apErrorFunction(err, "getAllOptions-Priority");
      });

    //Type of Deliverable Choices
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .fields.getByInternalNameOrTitle("ProjectType")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          let curTOD = TODacronymsCollection.filter((TOD) => {
            return TOD.ShortName == choice;
          });
          let choiceText =
            curTOD.length > 0 ? choice + " - " + curTOD[0].Name : choice;
          if (choice != null) {
            if (
              apModalBoxDrpDwnOptns.todOptns.findIndex((todOptn) => {
                return todOptn.key == choice;
              }) == -1
            ) {
              apModalBoxDrpDwnOptns.todOptns.push({
                key: choice,
                text: choiceText,
              });
            }
          }
        });
      })
      .then(() => {
        //apModalBoxDrpDwnOptns.todOptns.sort(_sortFilterKeys);
      })
      .catch((err) => {
        apErrorFunction(err, "getAllOptions-Type of Deliverable");
      });
    //Year Choices
    for (let year = moment().year(); year <= moment().year() + 10; year++) {
      apModalBoxDrpDwnOptns.yearOptns.push({
        key: year,
        text: year,
      });
    }
    //Term Choices
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .fields.getByInternalNameOrTitle("Term")()
      .then((response) => {
        apModalBoxDrpDwnOptns.termOptns = [];
        ["1", "2", "3", "4"].forEach((choice) => {
          if (choice != null) {
            if (
              apModalBoxDrpDwnOptns.termOptns.findIndex((termOptn) => {
                return termOptn.key == choice;
              }) == -1
            ) {
              apModalBoxDrpDwnOptns.termOptns.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then(() => {
        setApMasterProducts(apMasterProductCollection);
        setApModalBoxDropDownOptions(apModalBoxDrpDwnOptns);
      })
      .catch((err) => {
        apErrorFunction(err, "getAllOptions-Term");
      });
  };
  const filterKeys = (items) => {
    items.forEach((item) => {
      if (
        apDrpDwnOptns.baOptns.findIndex((baOptn) => {
          return baOptn.key == item.BusinessArea;
        }) == -1 &&
        item.BusinessArea
      ) {
        apDrpDwnOptns.baOptns.push({
          key: item.BusinessArea,
          text: item.BusinessArea,
        });
      }

      if (
        apDrpDwnOptns.todOptns.findIndex((todOptn) => {
          return todOptn.key == item.ProjectType;
        }) == -1 &&
        item.ProjectType
      ) {
        apDrpDwnOptns.todOptns.push({
          key: item.ProjectType,
          text: item.ProjectType,
        });
      }
      if (
        apDrpDwnOptns.PriorityOptns.findIndex((PriorityOptn) => {
          return PriorityOptn.key == item.Priority;
        }) == -1 &&
        item.Priority
      ) {
        apDrpDwnOptns.PriorityOptns.push({
          key: item.Priority,
          text: item.Priority,
        });
      }
      if (
        apDrpDwnOptns.yearOptns.findIndex((year) => {
          return year.key == item.Year;
        }) == -1 &&
        item.Year
      ) {
        apDrpDwnOptns.yearOptns.push({
          key: item.Year,
          text: item.Year,
        });
      }

      if (
        apDrpDwnOptns.potOptns.findIndex((potOptn) => {
          return potOptn.key == item.Title;
        }) == -1 &&
        item.Title
      ) {
        apDrpDwnOptns.potOptns.push({
          key: item.Title,
          text: item.Title,
        });
        apModalBoxDrpDwnOptns.potOptns.push({
          key: item.Title,
          text: item.Title,
        });
      }

      let tempmanager =
        item.ProjectOwnerId != null ? item.ProjectOwner.Title : null;
      if (
        apDrpDwnOptns.managerOptns.findIndex((managerOptn) => {
          return managerOptn.key == tempmanager;
        }) == -1 &&
        tempmanager
      ) {
        apDrpDwnOptns.managerOptns.push({
          key: tempmanager,
          text: tempmanager,
        });
      }

      let tempdevelopers = [];
      if (item.ProjectLeadId != null) {
        item.ProjectLead.forEach((dev) => {
          tempdevelopers.push(dev.Title);
        });

        tempdevelopers.forEach((tempdev) => {
          if (
            apDrpDwnOptns.developerOptns.findIndex((developerOptn) => {
              return developerOptn.key == tempdev;
            }) == -1 &&
            tempdev
          ) {
            apDrpDwnOptns.developerOptns.push({
              key: tempdev,
              text: tempdev,
            });
          }
        });
      }

      // if (
      //   apDrpDwnOptns.termOptns.findIndex((termOptn) => {
      //     return termOptn.key == item.Term;
      //   }) == -1 &&
      //   item.Term
      // ) {
      //   apDrpDwnOptns.termOptns.push({
      //     key: item.Term,
      //     text: item.Term,
      //   });
      // }
    });
    ["1", "2", "3", "4"].forEach((_item) => {
      if (
        apDrpDwnOptns.termOptns.findIndex((termOptn) => {
          return termOptn.key == _item;
        }) == -1 &&
        _item
      ) {
        apDrpDwnOptns.termOptns.push({
          key: _item,
          text: _item,
        });
      }
    });

    sortingFilterKeys(apDrpDwnOptns, apModalBoxDrpDwnOptns);

    setApDropDownOptions(apDrpDwnOptns);
    setApModalBoxDropDownOptions(apModalBoxDrpDwnOptns);
  };
  const filterKeysAfterModified = (items) => {
    items.forEach((item) => {
      if (
        apDrpDwnOptns.baOptns.findIndex((baOptn) => {
          return baOptn.key == item.BusinessArea;
        }) == -1 &&
        item.BusinessArea
      ) {
        apDrpDwnOptns.baOptns.push({
          key: item.BusinessArea,
          text: item.BusinessArea,
        });
      }

      if (
        apDrpDwnOptns.todOptns.findIndex((todOptn) => {
          return todOptn.key == item.TypeOfProject;
        }) == -1 &&
        item.TypeOfProject
      ) {
        apDrpDwnOptns.todOptns.push({
          key: item.TypeOfProject,
          text: item.TypeOfProject,
        });
      }
      if (
        apDrpDwnOptns.yearOptns.findIndex((year) => {
          return year.key == item.Year;
        }) == -1 &&
        item.Year
      ) {
        apDrpDwnOptns.yearOptns.push({
          key: item.Year,
          text: item.Year,
        });
      }
      if (
        apDrpDwnOptns.potOptns.findIndex((potOptn) => {
          return potOptn.key == item.ProjectOrTask;
        }) == -1 &&
        item.ProjectOrTask
      ) {
        apDrpDwnOptns.potOptns.push({
          key: item.ProjectOrTask,
          text: item.ProjectOrTask,
        });
        apModalBoxDrpDwnOptns.potOptns.push({
          key: item.ProjectOrTask,
          text: item.ProjectOrTask,
        });
      }

      let tempmanager = item.PMName != null ? item.PMName.name : null;
      if (
        apDrpDwnOptns.managerOptns.findIndex((managerOptn) => {
          return managerOptn.key == tempmanager;
        }) == -1 &&
        tempmanager
      ) {
        apDrpDwnOptns.managerOptns.push({
          key: tempmanager,
          text: tempmanager,
        });
      }

      let tempdevelopers = [];
      if (item.DNames.length > 0) {
        item.DNames.forEach((dev) => {
          tempdevelopers.push(dev.name);
        });

        tempdevelopers.forEach((tempdev) => {
          if (
            apDrpDwnOptns.developerOptns.findIndex((developerOptn) => {
              return developerOptn.key == tempdev;
            }) == -1 &&
            tempdev != null
          ) {
            apDrpDwnOptns.developerOptns.push({
              key: tempdev,
              text: tempdev,
            });
          }
        });
      }

      // if (
      //   apDrpDwnOptns.termOptns.findIndex((termOptn) => {
      //     return termOptn.key == item.Term;
      //   }) == -1 &&
      //   item.Term
      // ) {
      //   apDrpDwnOptns.termOptns.push({
      //     key: item.Term,
      //     text: item.Term,
      //   });
      // }
    });
    ["1", "2", "3", "4"].forEach((_item) => {
      if (
        apDrpDwnOptns.termOptns.findIndex((termOptn) => {
          return termOptn.key == _item;
        }) == -1 &&
        _item
      ) {
        apDrpDwnOptns.termOptns.push({
          key: _item,
          text: _item,
        });
      }
    });

    sortingFilterKeys(apDrpDwnOptns, apModalBoxDrpDwnOptns);

    setApDropDownOptions(apDrpDwnOptns);
    let tempArr = apModalBoxDropDownOptions;
    tempArr.potOptns = apModalBoxDrpDwnOptns.potOptns;
    setApModalBoxDropDownOptions(tempArr);
  };
  const sortingFilterKeys = (apDrpDwnOptns, apModalBoxDrpDwnOptns) => {
    const sortFilterKeys = (a, b) => {
      if (a.text < b.text) {
        return -1;
      }
      if (a.text > b.text) {
        return 1;
      }
      return 0;
    };

    apDrpDwnOptns.baOptns.shift();
    apDrpDwnOptns.baOptns.sort(sortFilterKeys);
    apDrpDwnOptns.baOptns.unshift({ key: "All", text: "All" });

    apDrpDwnOptns.todOptns.shift();
    apDrpDwnOptns.todOptns.sort(sortFilterKeys);
    apDrpDwnOptns.todOptns.unshift({ key: "All", text: "All" });

    apDrpDwnOptns.potOptns.shift();
    apDrpDwnOptns.potOptns.sort(sortFilterKeys);
    apDrpDwnOptns.potOptns.unshift({ key: "All", text: "All" });

    apModalBoxDrpDwnOptns.potOptns.sort(sortFilterKeys);

    if (
      apDrpDwnOptns.managerOptns.some((managerOptn) => {
        return (
          managerOptn.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      })
    ) {
      apDrpDwnOptns.managerOptns.shift();
      let loginUserIndex = apDrpDwnOptns.managerOptns.findIndex((user) => {
        return (
          user.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      });
      let loginUserData = apDrpDwnOptns.managerOptns.splice(loginUserIndex, 1);

      apDrpDwnOptns.managerOptns.sort(sortFilterKeys);
      apDrpDwnOptns.managerOptns.unshift(loginUserData[0]);
      apDrpDwnOptns.managerOptns = usersOrderFunction(
        apDrpDwnOptns.managerOptns
      );
      apDrpDwnOptns.managerOptns.unshift({ key: "All", text: "All" });
    } else {
      apDrpDwnOptns.managerOptns.shift();
      apDrpDwnOptns.managerOptns.sort(sortFilterKeys);
      apDrpDwnOptns.managerOptns = usersOrderFunction(
        apDrpDwnOptns.managerOptns
      );
      apDrpDwnOptns.managerOptns.unshift({ key: "All", text: "All" });
    }

    if (
      apDrpDwnOptns.developerOptns.some((developerOptn) => {
        return (
          developerOptn.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      })
    ) {
      apDrpDwnOptns.developerOptns.shift();
      let loginUserIndex = apDrpDwnOptns.developerOptns.findIndex((user) => {
        return (
          user.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      });
      let loginUserData = apDrpDwnOptns.developerOptns.splice(
        loginUserIndex,
        1
      );
      apDrpDwnOptns.developerOptns.sort(sortFilterKeys);
      apDrpDwnOptns.developerOptns.unshift(loginUserData[0]);
      apDrpDwnOptns.developerOptns = usersOrderFunction(
        apDrpDwnOptns.developerOptns
      );
      apDrpDwnOptns.developerOptns.unshift({ key: "All", text: "All" });
    } else {
      apDrpDwnOptns.developerOptns.shift();
      apDrpDwnOptns.developerOptns.sort(sortFilterKeys);
      apDrpDwnOptns.developerOptns = usersOrderFunction(
        apDrpDwnOptns.developerOptns
      );
      apDrpDwnOptns.developerOptns.unshift({ key: "All", text: "All" });
    }

    apDrpDwnOptns.termOptns.shift();
    apDrpDwnOptns.termOptns.sort(sortFilterKeys);
    apDrpDwnOptns.termOptns.unshift({ key: "All", text: "All" });
  };
  const usersOrderFunction = (dropDown): any => {
    let nonArchived = dropDown.filter((user) => {
      return !user.text.includes("Archive");
    });
    let archived = dropDown.filter((user) => {
      return user.text.includes("Archive");
    });

    return nonArchived.concat(archived);
  };
  const listFilter = (key, option) => {
    let tempArr = [...apMasterData];
    let tempApFilterKeys = { ...apFilterOptions };
    tempApFilterKeys[`${key}`] = option;

    if (tempApFilterKeys.ProjectOrTaskSearch) {
      tempArr = tempArr.filter((arr) => {
        return arr.ProjectOrTask.toLowerCase().includes(
          tempApFilterKeys.ProjectOrTaskSearch.toLowerCase()
        );
      });
    }
    if (tempApFilterKeys.BusinessArea != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.BusinessArea == tempApFilterKeys.BusinessArea;
      });
    }
    if (tempApFilterKeys.TypeOfProject != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.TypeOfProject == tempApFilterKeys.TypeOfProject;
      });
    }
    if (tempApFilterKeys.ProjectOrTask != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.ProjectOrTask == tempApFilterKeys.ProjectOrTask;
      });
    }
    if (tempApFilterKeys.PM != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.PMName.name == tempApFilterKeys.PM;
      });
    }
    if (tempApFilterKeys.Year != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Year == tempApFilterKeys.Year;
      });
    }
    if (tempApFilterKeys.D != "All") {
      let devArr = [];
      tempArr.forEach((arr) => {
        if (arr.DNames.length != 0) {
          if (arr.DNames.some((DName) => DName.name == tempApFilterKeys.D)) {
            devArr.push(arr);
          }
        }
      });
      tempArr = [...devArr];
    }
    if (tempApFilterKeys.Term != "All") {
      let termArr = [];
      tempArr.forEach((arr) => {
        if (arr.Term.length != 0) {
          if (arr.Term.some((term) => term == tempApFilterKeys.Term)) {
            termArr.push(arr);
          }
        }
      });
      tempArr = [...termArr];
    }

    filterKeysAfterModified(tempArr);
    paginatewithdata(1, tempArr);
    setApFilterOptions({ ...tempApFilterKeys });
    columnSortArr = tempArr;
    setApData(tempArr);
  };
  const listFilterAfterUpdated = (masterData: any) => {
    let tempArr = [...masterData];
    let tempApFilterKeys = { ...apFilterOptions };

    if (tempApFilterKeys.ProjectOrTaskSearch) {
      tempArr = tempArr.filter((arr) => {
        return arr.ProjectOrTask.toLowerCase().includes(
          tempApFilterKeys.ProjectOrTaskSearch.toLowerCase()
        );
      });
    }
    if (tempApFilterKeys.BusinessArea != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.BusinessArea == tempApFilterKeys.BusinessArea;
      });
    }
    if (tempApFilterKeys.TypeOfProject != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.TypeOfProject == tempApFilterKeys.TypeOfProject;
      });
    }
    if (tempApFilterKeys.ProjectOrTask != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.ProjectOrTask == tempApFilterKeys.ProjectOrTask;
      });
    }
    if (tempApFilterKeys.PM != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.PMName.name == tempApFilterKeys.PM;
      });
    }
    if (tempApFilterKeys.Year != "All") {
      tempArr = tempArr.filter((arr) => {
        return arr.Year == tempApFilterKeys.Year;
      });
    }
    if (tempApFilterKeys.D != "All") {
      let devArr = [];
      tempArr.forEach((arr) => {
        if (arr.DNames.length != 0) {
          if (arr.DNames.some((DName) => DName.name == tempApFilterKeys.D)) {
            devArr.push(arr);
          }
        }
      });
      tempArr = [...devArr];
    }
    if (tempApFilterKeys.Term != "All") {
      let termArr = [];
      tempArr.forEach((arr) => {
        if (arr.Term.length != 0) {
          if (arr.Term.some((term) => term == tempApFilterKeys.Term)) {
            termArr.push(arr);
          }
        }
      });
      tempArr = [...termArr];
    }

    return tempArr;
  };
  const dateFormater = (date: Date): string => {
    return date ? moment(date).format("DD/MM/YYYY") : "";
  };
  const GetUserDetails = (filterText, currentPersonas) => {
    let _allPeoples = allPeoples;

    _allPeoples = _allPeoples.filter((arr) => {
      return arr.text.toLowerCase().indexOf("archive") == -1;
    });

    if (currentPersonas.length > 0) {
      _allPeoples = _allPeoples.filter(
        (arr) => !currentPersonas.some((persona) => persona.ID == arr.ID)
      );
    }

    var result = _allPeoples.filter(
      (value, index, self) => index === self.findIndex((t) => t.ID === value.ID)
    );

    return result.filter((item) =>
      doesTextStartWith(item.text as string, filterText)
    );
  };

  const doesTextStartWith = (text: string, filterText: string) => {
    return text.toLowerCase().indexOf(filterText.toLowerCase()) === 0;
  };
  const onChangeHandler = (key: string, value: any) => {
    let tempResponseData = { ...apResponseData };
    if (key == "term") {
      if (value) {
        let tempResponseData = { ...apResponseData };
        tempResponseData.term = value.selected
          ? [...tempResponseData.term, value.key as string]
          : tempResponseData.term.filter((key) => key !== value.key);
        tempResponseData.term.sort();
        setApResponseData({ ...tempResponseData });
      }
    } else {
      if (key == "projectOrTask") {
        let curDeliverable = apMasterData.filter((arr) => {
          return arr.ID == tempResponseData.ID;
        })[0];

        let tempArr = apMasterData.filter((arr) => {
          return (
            arr.ProjectOrTask.toLowerCase() ==
            (value ? value.toLowerCase() : "")
          );
        });

        if (
          tempResponseData.ID == null ||
          curDeliverable.ProjectOrTask != value
        ) {
          tempResponseData["ProjectVersion"] = "V" + (tempArr.length + 1);
        } else {
          tempResponseData["ProjectVersion"] = curDeliverable.ProjectVersion;
        }
      }

      tempResponseData[key] = value;
      setApResponseData({ ...tempResponseData });
    }
  };
  const apAddItem = () => {
    let product = [];
    let devIds = [];
    if (apResponseData.developer.length > 0) {
      apResponseData.developer.forEach((dev) => {
        devIds.push(dev.ID);
      });
    }

    if (apResponseData.product != null) {
      product = apMasterProducts.filter((prod) => {
        return prod.ProductKey == apResponseData.product;
      });
    }

    const requestdata = {
      Title: apResponseData.projectOrTask ? apResponseData.projectOrTask : "",
      Status: "Scheduled",
      Master_x0020_ProjectId: product.length > 0 ? product[0].ProductId : null,
      ProjectOwnerId: apResponseData.manager ? apResponseData.manager : null,
      ProjectLeadId:
        apResponseData.developer.length > 0
          ? { results: [...devIds] }
          : { results: [] },
      Year: apResponseData.year ? apResponseData.year : null,
      TermNew:
        apResponseData.term.length > 0
          ? { results: [...apResponseData.term] }
          : { results: [] },
      BusinessArea: apResponseData.businessArea
        ? apResponseData.businessArea
        : null,
      Priority: apResponseData.Priority ? apResponseData.Priority : null,
      ProjectVersion: apResponseData.ProjectVersion
        ? apResponseData.ProjectVersion
        : "",
      BA_x0020_acronyms: apResponseData.businessArea
        ? BAacronymsCollection.filter((BAacronym) => {
            return BAacronym.Name == apResponseData.businessArea;
          })[0].ShortName
        : null,
      ProjectType: apResponseData.typeOfProject
        ? apResponseData.typeOfProject
        : null,
      StartDate: apResponseData.startDate
        ? moment(apResponseData.startDate, DateListFormat).format("YYYY-MM-DD")
        : null,
      PlannedEndDate: apResponseData.endDate
        ? moment(apResponseData.endDate, DateListFormat).format("YYYY-MM-DD")
        : null,
    };

    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.add(requestdata)
      .then((e) => {
        sharepointWeb.lists
          .getByTitle(ListNameURL)
          .items.getById(e.data.Id)
          .select(
            "*",
            "ProjectOwner/Title",
            "ProjectOwner/Id",
            "ProjectOwner/EMail",
            "ProjectLead/Title",
            "ProjectLead/Id",
            "ProjectLead/EMail",
            "Master_x0020_Project/Title",
            "Master_x0020_Project/Id",
            "Master_x0020_Project/ProductVersion",
            "FieldValuesAsText/StartDate",
            "FieldValuesAsText/PlannedEndDate"
          )
          .expand(
            "ProjectOwner",
            "ProjectLead",
            "Master_x0020_Project",
            "FieldValuesAsText"
          )
          .get()
          .then((item) => {
            let tempMasterArr = [...apMasterData];
            let newItemAddedtoArr = [];
            let arrAfterAddApData = allitemsArrayFormatter(
              item,
              newItemAddedtoArr
            );

            Array.prototype.push.apply(arrAfterAddApData, tempMasterArr);

            filterKeysAfterModified(arrAfterAddApData);
            let lastIndex: number = 1 * totalPageItems;
            let firstIndex: number = lastIndex - totalPageItems;
            let paginatedItems = arrAfterAddApData.slice(firstIndex, lastIndex);

            setApModalBoxVisibility({
              condition: false,
              action: "",
              selectedItem: [],
            });

            setApUnsortMasterData([...arrAfterAddApData]);
            columnSortArr = arrAfterAddApData;
            setApData(arrAfterAddApData);
            columnSortMasterArr = arrAfterAddApData;
            setApMasterData([...arrAfterAddApData]);
            setdisplayData([...paginatedItems]);
            setApCurrentPage(1);
            setApShowMessage(apErrorStatus);
            setApResponseData({ ...responseData });
            setApModelBoxDrpDwnToTxtBox(false);
            setApOnSubmitLoader(false);
            AddSuccessPopup();
          })
          .catch((err) => {
            apErrorFunction(err, "apAddItem-getItem");
          });
      })
      .catch((err) => {
        apErrorFunction(err, "apAddItem-updateItem");
      });
  };
  const apUpdateItem = (id: number) => {
    let product = [];
    let devIds = [];
    if (apResponseData.developer.length > 0) {
      apResponseData.developer.forEach((dev) => {
        devIds.push(dev.ID);
      });
    }

    if (apResponseData.product != null) {
      product = apMasterProducts.filter((prod) => {
        return prod.ProductKey == apResponseData.product;
      });
    }

    const requestdata = {
      Title: apResponseData.projectOrTask ? apResponseData.projectOrTask : "",
      Master_x0020_ProjectId: product.length > 0 ? product[0].ProductId : null,
      ProjectOwnerId: apResponseData.manager ? apResponseData.manager : null,
      ProjectLeadId:
        apResponseData.developer.length > 0
          ? { results: [...devIds] }
          : { results: [] },
      Year: apResponseData.year ? apResponseData.year : null,
      TermNew:
        apResponseData.term.length > 0
          ? { results: [...apResponseData.term] }
          : { results: [] },
      BusinessArea: apResponseData.businessArea
        ? apResponseData.businessArea
        : null,
      BA_x0020_acronyms: apResponseData.businessArea
        ? BAacronymsCollection.filter((BAacronym) => {
            return BAacronym.Name == apResponseData.businessArea;
          })[0].ShortName
        : null,
      ProjectType: apResponseData.typeOfProject
        ? apResponseData.typeOfProject
        : null,
      Priority: apResponseData.Priority ? apResponseData.Priority : null,
      ProjectVersion: apResponseData.ProjectVersion
        ? apResponseData.ProjectVersion
        : "",
      StartDate: apResponseData.startDate
        ? moment(apResponseData.startDate, DateListFormat).format("YYYY-MM-DD")
        : null,
      PlannedEndDate: apResponseData.endDate
        ? moment(apResponseData.endDate, DateListFormat).format("YYYY-MM-DD")
        : null,
      Status:
        apResponseData.status == "Completed"
          ? "Completed"
          : apResponseData.status == "On hold"
          ? "On hold"
          : apResponseData.status,
    };

    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.getById(id)
      .update(requestdata)
      .then(() => {
        sharepointWeb.lists
          .getByTitle(ListNameURL)
          .items.getById(id)
          .select(
            "*",
            "ProjectOwner/Title",
            "ProjectOwner/Id",
            "ProjectOwner/EMail",
            "ProjectLead/Title",
            "ProjectLead/Id",
            "ProjectLead/EMail",
            "Master_x0020_Project/Title",
            "Master_x0020_Project/Id",
            "Master_x0020_Project/ProductVersion",
            "FieldValuesAsText/StartDate",
            "FieldValuesAsText/PlannedEndDate"
          )
          .expand(
            "ProjectOwner",
            "ProjectLead",
            "Master_x0020_Project",
            "FieldValuesAsText"
          )
          .get()
          .then((item) => {
            let tempMasterArr = [...apMasterData];
            let updatedItemtoArr = [];
            let targetIndex = tempMasterArr.findIndex((arr) => arr.ID == id);
            tempMasterArr.splice(targetIndex, 1);
            let arrAfterUpdateApData = allitemsArrayFormatter(
              item,
              updatedItemtoArr
            );
            Array.prototype.push.apply(arrAfterUpdateApData, tempMasterArr);

            let ArrAfterUpdated = listFilterAfterUpdated(arrAfterUpdateApData);

            filterKeysAfterModified(ArrAfterUpdated);

            let lastIndex: number = 1 * totalPageItems;
            let firstIndex: number = lastIndex - totalPageItems;
            let paginatedItems = ArrAfterUpdated.slice(firstIndex, lastIndex);

            setApModalBoxVisibility({
              condition: false,
              action: "",
              selectedItem: [],
            });

            setApUnsortMasterData(arrAfterUpdateApData);
            columnSortMasterArr = arrAfterUpdateApData;
            setApMasterData([...arrAfterUpdateApData]);
            columnSortArr = ArrAfterUpdated;
            setApData(ArrAfterUpdated);
            setdisplayData([...paginatedItems]);
            setApCurrentPage(1);
            setApShowMessage(apErrorStatus);
            setApResponseData({ ...responseData });
            setApOnSubmitLoader(false);
            setApSubmitConfirmLoader(false);
            setSubmitConfirmationPopup(false);
            UpdateSuccessPopup();
          })
          .catch((err) => {
            apErrorFunction(err, "apUpdateItem-updateItem");
          });
      })
      .catch((err) => {
        apErrorFunction(err, "apUpdateItem-updateItem");
      });
  };
  const apDeleteItem = (id: number) => {
    sharepointWeb.lists
      .getByTitle(ListNameURL)
      .items.getById(id)
      .delete()
      .then(() => {
        let tempMasterArr = [...apMasterData];
        let targetIndex = tempMasterArr.findIndex((arr) => arr.ID == id);
        tempMasterArr.splice(targetIndex, 1);

        let temp_ap_arr = [...apData];
        let targetIndexapdata = temp_ap_arr.findIndex((arr) => arr.ID == id);
        temp_ap_arr.splice(targetIndexapdata, 1);

        filterKeysAfterModified(temp_ap_arr);

        setApUnsortMasterData(tempMasterArr);
        columnSortMasterArr = tempMasterArr;
        setApMasterData(tempMasterArr);
        columnSortArr = temp_ap_arr;
        setApData(temp_ap_arr);
        paginatewithdata(apcurrentPage, temp_ap_arr);
        setApOnDeleteLoader(false);
        setApDeletePopup({ condition: false, targetId: 0 });
        DeleteSuccessPopup();
      })
      .catch((err) => {
        apErrorFunction(err, "apDeleteItem");
      });
  };
  const apValidationFunction = () => {
    let isError = false;

    let errorStatus = {
      businessAreaError: "",
      projectOrTaskError: "",
      productError: "",
    };
    if (!apResponseData.product) {
      isError = true;
      errorStatus.productError = "Please select product or solution";
    }

    if (!apResponseData.businessArea) {
      isError = true;
      errorStatus.businessAreaError = "Please select business area";
    }
    if (!apResponseData.projectOrTask) {
      isError = true;
      errorStatus.projectOrTaskError = "Please select name of the deliverable";
    }

    if (!isError) {
      if (apModalBoxVisibility.action == "Add") {
        setApOnSubmitLoader(true);
        apAddItem();
      } else if (apModalBoxVisibility.action == "Update") {
        let filteredArr = apMasterData.filter((data) => {
          return data.ID == apResponseData.ID;
        })[0];
        if (
          apResponseData.status == "Completed" &&
          filteredArr.Status != "Completed"
        ) {
          setSubmitConfirmationPopup(true);
        } else {
          setApOnSubmitLoader(true);
          apUpdateItem(apResponseData.ID);
        }
      }
    } else {
      setApShowMessage(errorStatus);
    }
  };
  const paginate = (pagenumber) => {
    let lastIndex: number = pagenumber * totalPageItems;
    let firstIndex: number = lastIndex - totalPageItems;
    let paginatedItems = apData.slice(firstIndex, lastIndex);
    currentpage = pagenumber;
    setdisplayData(paginatedItems);
    setApCurrentPage(pagenumber);
  };
  const paginatewithdata = (pagenumber, data) => {
    let lastIndex: number = pagenumber * totalPageItems;
    let firstIndex: number = lastIndex - totalPageItems;
    let paginatedItems = data.slice(firstIndex, lastIndex);
    currentpage = pagenumber;
    if (paginatedItems.length > 0) {
      setdisplayData(paginatedItems);
      setApCurrentPage(pagenumber);
    } else {
      paginate(pagenumber - 1);
    }
  };
  const generateExcel = () => {
    let arrExport = apData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "Business area", key: "businessArea", width: 25 },
      { header: "Term", key: "term", width: 25 },
      { header: "Hours", key: "hours", width: 25 },
      { header: "Start date", key: "startDate", width: 25 },
      { header: "End date", key: "endDate", width: 25 },
      { header: "Product or solution", key: "product", width: 60 },
      { header: "Type of deliverbale", key: "typeOfDeliverable", width: 20 },
      { header: "Name of the deliverable", key: "projectOrTask", width: 40 },
      { header: "Priority", key: "Priority", width: 60 },
      { header: "Status", key: "status", width: 30 },
      { header: "Client", key: "manager", width: 30 },
      { header: "Developer", key: "developer", width: 30 },
    ];
    arrExport.forEach((item) => {
      worksheet.addRow({
        businessArea: item.BusinessArea ? item.BusinessArea : "",
        term: item.Term ? parseInt(item.Term) : "",
        hours: item.Hours ? item.Hours : "",
        startDate: item.StartDate ? item.StartDate : "",
        endDate: item.EndDate ? item.EndDate : "",
        product: item.Product ? item.Product : "",
        typeOfDeliverable: item.TypeOfProject ? item.TypeOfProject : "",
        projectOrTask: item.ProjectOrTask ? item.ProjectOrTask : "",
        status: item.StatusStage ? item.StatusStage : "",
        manager: item.PMName ? item.PMName.name : "",
        Priority: item.Priority ? item.Priority : "",
        developer: item.DNames.length > 0 ? item.DNames[0].name : "",
      });
    });
    [
      "A1",
      "B1",
      "C1",
      "D1",
      "E1",
      "F1",
      "G1",
      "H1",
      "I1",
      "J1",
      "K1",
      "L1",
    ].map((key) => {
      worksheet.getCell(key).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "00e8d1" },
      };
    });
    [
      "A1",
      "B1",
      "C1",
      "D1",
      "E1",
      "F1",
      "G1",
      "H1",
      "I1",
      "J1",
      "K1",
      "L1",
    ].map((key) => {
      worksheet.getCell(key).color = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF" },
      };
    });
    workbook.xlsx
      .writeBuffer()
      .then((buffer) =>
        FileSaver.saveAs(
          new Blob([buffer]),
          `AnnualPlan-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
  };
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempapColumns = apColumns;
    const newColumns: IColumn[] = tempapColumns.slice();
    const currColumn: IColumn = newColumns.filter(
      (currCol) => column.key === currCol.key
    )[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });

    const newData = _copyAndSort(
      columnSortArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newMasterData = _copyAndSort(
      columnSortMasterArr,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setApData([...newData]);
    setApMasterData([...newMasterData]);
    paginatewithdata(1, newData);
  };
  function _copyAndSort<T>(
    items: T[],
    columnKey: string,
    isSortedDescending?: boolean
  ): T[] {
    let key = columnKey as keyof T;
    if (key == "PM") {
      const ascSortFunction = (a, b) => {
        if (a.PMName["name"] < b.PMName["name"]) {
          return -1;
        }
        if (a.PMName["name"] > b.PMName["name"]) {
          return 1;
        }
        return 0;
      };
      const decSortFunction = (b, a) => {
        if (a.PMName["name"] < b.PMName["name"]) {
          return -1;
        }
        if (a.PMName["name"] > b.PMName["name"]) {
          return 1;
        }
        return 0;
      };

      return items.sort(isSortedDescending ? ascSortFunction : decSortFunction);
    } else {
      return items
        .slice(0)
        .sort((a: T, b: T) =>
          (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
        );
    }
  }
  const AddSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Annual plan is successfully submitted !!!")
  );
  const UpdateSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Annual plan is successfully updated !!!")
  );
  const DeleteSuccessPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.success("Annual plan is successfully deleted !!!")
  );
  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );
  const apErrorFunction = (error: any, functionName: string) => {
    console.log(error);

    let response = {
      ComponentName: "Annual plan",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: loggeduseremail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setApStartUpLoader(false);
        setApOnSubmitLoader(false);
        setApOnDeleteLoader(false);
        ErrorPopup();
      }
    );
  };

  useEffect(() => {
    getAllOptions();
    getApData();
  }, [apReRender]);

  return (
    <div style={{ padding: "5px 15px" }}>
      {apStartUpLoader ? <CustomLoader /> : null}
      <div className={styles.apHeaderSection}>
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
          }}
        >
          <div className={styles.apHeader}>Annual plan</div>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <div style={{ display: "flex", alignItems: "center" }}>
              <Label styles={apLabelStyles}>
                Number of records :{" "}
                <b style={{ color: "#038387" }}>{apData.length}</b>
              </Label>
              <Label
                onClick={() => {
                  generateExcel();
                }}
                style={{
                  backgroundColor: "#EBEBEB",
                  padding: "0 15px",
                  cursor: "pointer",
                  fontSize: "12px",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  borderRadius: "3px",
                  color: "#1D6F42",
                  height: 34,
                }}
              >
                <Icon
                  style={{
                    color: "#1D6F42",
                  }}
                  iconName="ExcelDocument"
                  className={apIconStyleClass.export}
                />
                Export as XLS
              </Label>
              <div className={styles.apAddBtn}>
                <a
                  onClick={(_) => {
                    columnSortArr = apUnsortMasterData;
                    setApData(apUnsortMasterData);
                    columnSortMasterArr = apUnsortMasterData;
                    setApMasterData(apUnsortMasterData);
                    filterKeysAfterModified(apMasterData);
                    setApFilterOptions({ ...apFilterKeys });
                    paginatewithdata(1, apUnsortMasterData);
                    setApModalBoxVisibility({
                      condition: true,
                      action: "Add",
                      selectedItem: [],
                    });
                  }}
                >
                  Add deliverable
                </a>
              </div>
              {props.isAdmin ? (
                <div className={styles.apAddBtn}>
                  <a
                    onClick={(_) => {
                      props.handleclick("MasterProduct");
                    }}
                  >
                    Products
                  </a>
                </div>
              ) : null}
            </div>
          </div>
        </div>
        {/* Dropdown Section */}
        <div style={{ display: "flex", flexWrap: "wrap" }}>
          <div>
            <Label styles={apLabelStyles}>Search</Label>
            <SearchBox
              placeholder="Find deliverable"
              styles={
                apFilterOptions.ProjectOrTaskSearch == ""
                  ? apSearchBoxStyles
                  : apActiveSearchBoxStyles
              }
              value={apFilterOptions.ProjectOrTaskSearch}
              onChange={(e, value) => {
                listFilter("ProjectOrTaskSearch", value);
              }}
            />
          </div>
          <div>
            <Label styles={apLabelStyles}>Business area</Label>
            <Dropdown
              placeholder="Select a business area"
              options={apDropDownOptions.baOptns}
              styles={
                apFilterOptions.BusinessArea == "All"
                  ? apDropdownStyles
                  : apActiveDropdownStyles
              }
              onChange={(e, option: any) => {
                listFilter("BusinessArea", option["key"]);
              }}
              selectedKey={apFilterOptions.BusinessArea}
            />
          </div>
          <div>
            <Label styles={apLabelStyles}>Type of deliverable</Label>
            <Dropdown
              selectedKey={apFilterOptions.TypeOfProject}
              placeholder="Select a type of deliverable"
              options={apDropDownOptions.todOptns}
              styles={
                apFilterOptions.TypeOfProject == "All"
                  ? apDropdownStyles
                  : apActiveDropdownStyles
              }
              onChange={(e, option: any) => {
                listFilter("TypeOfProject", option["key"]);
              }}
            />
          </div>
          <div>
            <Label styles={apLabelStyles}>Name of the deliverable</Label>
            <Dropdown
              selectedKey={apFilterOptions.ProjectOrTask}
              placeholder="Select a deliverable"
              options={apDropDownOptions.potOptns}
              dropdownWidth={"auto"}
              styles={
                apFilterOptions.ProjectOrTask == "All"
                  ? apDropdownStyles
                  : apActiveDropdownStyles
              }
              onChange={(e, option: any) => {
                listFilter("ProjectOrTask", option["key"]);
              }}
            />
          </div>
          <div>
            <Label styles={apLabelStyles}>Client</Label>
            <Dropdown
              selectedKey={apFilterOptions.PM}
              placeholder="Select client"
              options={apDropDownOptions.managerOptns}
              styles={
                apFilterOptions.PM == "All"
                  ? apDropdownStyles
                  : apActiveDropdownStyles
              }
              onChange={(e, option: any) => {
                listFilter("PM", option["key"]);
              }}
            />
          </div>
          <div>
            <Label styles={apLabelStyles}>Developer</Label>
            <Dropdown
              selectedKey={apFilterOptions.D}
              placeholder="Select developers"
              options={apDropDownOptions.developerOptns}
              styles={
                apFilterOptions.D == "All"
                  ? apDropdownStyles
                  : apActiveDropdownStyles
              }
              onChange={(e, option: any) => {
                listFilter("D", option["key"]);
              }}
            />
          </div>
          <div>
            <Label styles={apShortLabelStyles}>Term</Label>
            <Dropdown
              selectedKey={apFilterOptions.Term}
              multiSelect={false}
              placeholder="Select terms"
              options={apDropDownOptions.termOptns}
              styles={
                apFilterOptions.Term == "All"
                  ? apShortDropdownStyles
                  : apActiveShortDropdownStyles
              }
              onChange={(e, option: any) => {
                listFilter("Term", option["key"]);
              }}
            />
          </div>
          <div>
            <Label styles={apShortLabelStyles}>Year</Label>
            <Dropdown
              selectedKey={apFilterOptions.Year}
              multiSelect={false}
              placeholder="Select year"
              options={apDropDownOptions.yearOptns}
              styles={
                apFilterOptions.Year == "All"
                  ? apShortDropdownStyles
                  : apActiveShortDropdownStyles
              }
              onChange={(e, option: any) => {
                listFilter("Year", option["key"]);
              }}
            />
          </div>
          <div>
            <div>
              <Icon
                iconName="Refresh"
                title="Click to reset"
                className={apIconStyleClass.refresh}
                onClick={() => {
                  columnSortArr = apUnsortMasterData;
                  setApData([...apUnsortMasterData]);
                  columnSortMasterArr = apUnsortMasterData;
                  setApMasterData([...apUnsortMasterData]);
                  setMasterApColumn(apColumns);
                  filterKeysAfterModified(apMasterData);
                  setApFilterOptions({ ...apFilterKeys });
                  paginatewithdata(1, apUnsortMasterData);
                }}
              />
            </div>
          </div>
        </div>
        {/* Dropdown Section */}
      </div>

      <div>
        <DetailsList
          items={displayData}
          columns={masterApColumn}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
          styles={gridStyles}
        />
      </div>
      <div>
        {displayData.length > 0 ? (
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              margin: "10px 0",
            }}
          >
            <Pagination
              currentPage={apcurrentPage}
              totalPages={
                apData.length > 0
                  ? Math.ceil(apData.length / totalPageItems)
                  : 1
              }
              onChange={(page) => {
                paginate(page);
              }}
            />
          </div>
        ) : (
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              marginTop: "15px",
            }}
          >
            <Label style={{ color: "#2392B2" }}>No data Found !!!</Label>
          </div>
        )}
      </div>
      {/* Add & Update Popup */}
      <div>
        {apModalBoxVisibility.condition ? (
          <Modal isOpen={apModalBoxVisibility.condition} isBlocking={true}>
            <div style={{ padding: "30px 20px" }}>
              <div
                style={{
                  fontSize: 24,
                  textAlign: "center",
                  color: "#2392B2",
                  fontWeight: "600",
                  marginBottom: "20px",
                }}
              >
                {apModalBoxVisibility.action == "Add"
                  ? "New deliverable "
                  : apModalBoxVisibility.action == "Update"
                  ? "Edit deliverable"
                  : ""}
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  justifyContent: "flex-start",
                }}
              >
                <div>
                  <Dropdown
                    label="Business area"
                    required={true}
                    errorMessage={apShowMessage.businessAreaError}
                    selectedKey={apResponseData.businessArea}
                    placeholder="Select a business area"
                    options={apModalBoxDropDownOptions.baOptns}
                    styles={apModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      onChangeHandler("businessArea", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Dropdown
                    label="Type of deliverable"
                    selectedKey={apResponseData.typeOfProject}
                    placeholder="Select a type of deliverable"
                    options={apModalBoxDropDownOptions.todOptns}
                    styles={apModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      onChangeHandler("typeOfProject", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <Dropdown
                    label="Term"
                    selectedKeys={
                      apResponseData.term.length > 0
                        ? [...apResponseData.term]
                        : []
                    }
                    multiSelect
                    placeholder="Select terms"
                    options={apModalBoxDropDownOptions.termOptns}
                    styles={apModalBoxDropdownStyles}
                    onChange={(
                      event: React.FormEvent<HTMLDivElement>,
                      item: IDropdownOption
                    ): void => {
                      onChangeHandler("term", item);
                      // if (item) {
                      //   let tempResponseData = { ...apResponseData };
                      //   tempResponseData.term = item.selected
                      //     ? [...tempResponseData.term, item.key as string]
                      //     : tempResponseData.term.filter(
                      //         (key) => key !== item.key
                      //       );
                      //   tempResponseData.term.sort();
                      //   setApResponseData({ ...tempResponseData });
                      // }
                    }}
                  />
                </div>
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  justifyContent: "flex-start",
                }}
              >
                <div>
                  <Dropdown
                    label="Product or solution"
                    selectedKey={apResponseData.product}
                    required={true}
                    errorMessage={apShowMessage.productError}
                    placeholder="Select a product or solution"
                    options={apModalBoxDropDownOptions.productOptns}
                    styles={apModalBoxDrpDwnCalloutStyles}
                    onChange={(e, option: any) => {
                      onChangeHandler("product", option["key"]);
                    }}
                  />
                </div>
                <div>
                  <DatePicker
                    label="Start date"
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                    formatDate={dateFormater}
                    value={apResponseData.startDate}
                    styles={apModalBoxDatePickerStyles}
                    onSelectDate={(value: any) => {
                      onChangeHandler("startDate", value);
                    }}
                  />
                </div>
                <div>
                  <DatePicker
                    label="End date"
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                    styles={apModalBoxDatePickerStyles}
                    formatDate={dateFormater}
                    value={apResponseData.endDate}
                    onSelectDate={(value: any) => {
                      onChangeHandler("endDate", value);
                    }}
                  />
                </div>
              </div>
              {apModalBoxVisibility.action == "Add" ? (
                <div
                  style={{
                    display: "flex",
                    alignItems: "flex-start",
                    justifyContent: "space-between",
                  }}
                >
                  <div>
                    {apModelBoxDrpDwnToTxtBox ? (
                      <TextField
                        label="Name of the deliverable"
                        placeholder="Add name of the deliverable"
                        errorMessage={apShowMessage.projectOrTaskError}
                        required={true}
                        defaultValue={apResponseData.projectOrTask}
                        styles={apTxtBoxStyles}
                        onChange={(e, value: string) => {
                          onChangeHandler("projectOrTask", value);
                        }}
                      />
                    ) : (
                      <Dropdown
                        label="Name of the deliverable"
                        selectedKey={apResponseData.projectOrTask}
                        placeholder="Select name of the deliverable"
                        errorMessage={apShowMessage.projectOrTaskError}
                        required={true}
                        options={apModalBoxDropDownOptions.potOptns}
                        styles={apModalBoxDrpDwnCalloutStyles}
                        style={{ width: "780px" }}
                        onChange={(e, option: any) => {
                          onChangeHandler("projectOrTask", option["key"]);
                        }}
                      />
                    )}
                  </div>
                  <div
                    style={{
                      width: "21.5%",
                      display: "flex",
                      justifyContent: "space-between",
                    }}
                  >
                    <div
                      style={{
                        marginLeft: apModelBoxDrpDwnToTxtBox ? "-15px" : "0px",
                      }}
                    >
                      <TextField
                        label="Version"
                        placeholder=""
                        disabled
                        value={
                          apResponseData.ProjectVersion
                            ? apResponseData.ProjectVersion
                            : "V1"
                        }
                        styles={apTxtBoxStylesSmall}
                        onChange={(e, value: string) => {
                          // onChangeHandler("projectOrTask", value);
                        }}
                      />
                    </div>
                    <div>
                      {apModelBoxDrpDwnToTxtBox ? (
                        <Checkbox
                          label="New"
                          styles={apModalBoxCheckBoxStyles}
                          checked={apModelBoxDrpDwnToTxtBox}
                          onChange={(e) => {
                            onChangeHandler("projectOrTask", "");
                            setApModelBoxDrpDwnToTxtBox(
                              !apModelBoxDrpDwnToTxtBox
                            );
                          }}
                        />
                      ) : (
                        <Checkbox
                          label="New"
                          styles={apModalBoxCheckBoxStyles}
                          checked={apModelBoxDrpDwnToTxtBox}
                          onChange={(e) => {
                            onChangeHandler("projectOrTask", "");
                            setApModelBoxDrpDwnToTxtBox(
                              !apModelBoxDrpDwnToTxtBox
                            );
                          }}
                        />
                      )}
                    </div>
                  </div>
                </div>
              ) : apModalBoxVisibility.action == "Update" ? (
                <div
                  style={{
                    display: "flex",
                  }}
                >
                  <div>
                    <TextField
                      label="Name of the deliverable"
                      placeholder="Add name of the deliverable"
                      errorMessage={apShowMessage.projectOrTaskError}
                      defaultValue={apResponseData.projectOrTask}
                      required={true}
                      styles={apTxtBoxStyles}
                      onChange={(e, value: string) => {
                        onChangeHandler("projectOrTask", value);
                      }}
                    />
                  </div>
                  <div
                    style={{
                      marginLeft: "-15px",
                    }}
                  >
                    <TextField
                      label="Version"
                      placeholder=""
                      disabled
                      value={
                        apResponseData.ProjectVersion
                          ? apResponseData.ProjectVersion
                          : "V1"
                      }
                      styles={apTxtBoxStylesSmall}
                      onChange={(e, value: string) => {
                        // onChangeHandler("projectOrTask", value);
                      }}
                    />
                  </div>
                </div>
              ) : (
                ""
              )}
              <div
                style={{
                  display: "flex",
                  alignItems: "flex-start",
                  justifyContent: "flex-start",
                }}
              >
                <div
                  style={{
                    display: "flex",
                    alignItems: "flex-start",
                    justifyContent: "flex-start",
                    flexWrap: "wrap",
                    width: "680px",
                  }}
                >
                  <div>
                    <Dropdown
                      label="Year"
                      selectedKey={apResponseData.year}
                      placeholder="Select year"
                      options={
                        apModalBoxVisibility.action == "Update"
                          ? editYear
                          : apModalBoxDropDownOptions.yearOptns
                      }
                      styles={apModalBoxDrpDwnCalloutStyles}
                      onChange={(e, option: any) => {
                        onChangeHandler("year", option["key"]);
                      }}
                    />
                  </div>
                  <div>
                    <Label
                      style={{
                        transform: "translate(20px, 10px)",
                      }}
                    >
                      Client
                    </Label>
                    <NormalPeoplePicker
                      className={apModalBoxPP}
                      onResolveSuggestions={GetUserDetails}
                      itemLimit={1}
                      defaultSelectedItems={allPeoples.filter((people) => {
                        return people.ID == apResponseData.manager;
                      })}
                      styles={{
                        root: {
                          width: 300,
                          margin: "10px 20px",
                          selectors: {
                            ".ms-BasePicker-text": {
                              height: 36,
                              padding: "3px 10px",
                              border: "1px solid black",
                              borderRadius: 4,
                            },
                          },
                          ".ms-Persona-primaryText": { fontWeight: 600 },
                        },
                      }}
                      onChange={(selectedUser) => {
                        selectedUser.length != 0
                          ? onChangeHandler("manager", selectedUser[0]["ID"])
                          : onChangeHandler("manager", "");
                      }}
                    />
                  </div>
                  <div
                    style={{
                      display: "flex",
                      alignItems: "flex-start",
                      justifyContent: "flex-start",
                    }}
                  >
                    <Dropdown
                      label="Priority"
                      selectedKey={apResponseData.Priority}
                      // disabled={
                      //   apMasterData.filter((data) => {
                      //     return data.ID == apResponseData.ID;
                      //   })[0].Status == "Completed"
                      //     ? true
                      //     : false
                      // }
                      placeholder="Select Priority"
                      options={apModalBoxDropDownOptions.PriorityOptns}
                      styles={apModalBoxDropdownStyles}
                      onChange={(e, option: any) => {
                        onChangeHandler("Priority", option["key"]);
                      }}
                    />
                    {apModalBoxVisibility.action == "Update" ? (
                      <Dropdown
                        label="Status"
                        selectedKey={
                          apResponseData.status == "Completed" ||
                          apResponseData.status == "On hold"
                            ? apResponseData.status
                            : ""
                        }
                        disabled={
                          apMasterData.filter((data) => {
                            return data.ID == apResponseData.ID;
                          })[0].Status == "Completed"
                            ? true
                            : false
                        }
                        placeholder="Select status"
                        options={apModalBoxDropDownOptions.statusOtpns}
                        styles={apModalBoxDropdownStyles}
                        onChange={(e, option: any) => {
                          onChangeHandler("status", option["key"]);
                        }}
                      />
                    ) : (
                      ""
                    )}
                  </div>
                </div>
                <div>
                  <Label
                    style={{
                      transform: "translate(20px, 10px)",
                    }}
                  >
                    Developer
                  </Label>
                  <NormalPeoplePicker
                    className={apModalBoxPP}
                    onResolveSuggestions={GetUserDetails}
                    defaultSelectedItems={apResponseData.developer}
                    styles={{
                      root: {
                        width: 300,
                        margin: "10px 20px",
                        selectors: {
                          ".ms-BasePicker-text": {
                            padding: "3px 10px",
                            border: "1px solid black",
                            borderRadius: 4,
                            maxHeight: "115px",
                            overflowX: "hidden",
                            "::after": {
                              border: "none",
                            },
                          },
                        },
                        ".ms-Persona-primaryText": {
                          fontWeight: 600,
                          border: "none",
                        },
                      },
                    }}
                    onChange={(selectedUser) => {
                      selectedUser.length != 0
                        ? onChangeHandler("developer", selectedUser)
                        : onChangeHandler("developer", "");
                    }}
                  />
                </div>
              </div>
              <div style={{ display: "flex", justifyContent: "end" }}>
                <div className={styles.apModalBoxButtonSection}>
                  <button
                    className={styles.apModalBoxSubmitBtn}
                    onClick={(_) => {
                      apValidationFunction();
                    }}
                    style={{ display: "flex" }}
                  >
                    {apOnSubmitLoader ? (
                      <Spinner />
                    ) : apModalBoxVisibility.action == "Add" ? (
                      <span>
                        <Icon
                          iconName="Save"
                          style={{ position: "relative", top: 3, left: -8 }}
                        />
                        {"Add"}
                      </span>
                    ) : apModalBoxVisibility.action == "Update" ? (
                      <span>
                        <Icon
                          iconName="Save"
                          style={{ position: "relative", top: 3, left: -8 }}
                        />
                        {"Update"}
                      </span>
                    ) : (
                      ""
                    )}
                  </button>
                  <button
                    className={styles.apModalBoxBackBtn}
                    onClick={(_) => {
                      setApResponseData({ ...responseData });
                      setApShowMessage(apErrorStatus);
                      setApModelBoxDrpDwnToTxtBox(false);
                      setApModalBoxVisibility({
                        condition: false,
                        action: "",
                        selectedItem: [],
                      });
                    }}
                  >
                    <span>
                      {" "}
                      <Icon
                        iconName="Cancel"
                        style={{ position: "relative", top: 3, left: -8 }}
                      />
                      Close
                    </span>
                  </button>
                </div>
              </div>
            </div>
          </Modal>
        ) : (
          ""
        )}
      </div>
      {/* Delete Popup */}
      <div>
        {apDeletePopup.condition ? (
          <Modal isOpen={apDeletePopup.condition} isBlocking={true}>
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                alignItems: "center",
                marginTop: "30px",
                width: "450px",
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "flex-start",
                  flexDirection: "column",
                  marginBottom: "10px",
                }}
              >
                <Label className={styles.deletePopupTitle}>
                  Delete deliverable
                </Label>
                <Label
                  style={{
                    padding: "5px 20px",
                  }}
                  className={styles.deletePopupDesc}
                >
                  The data up to production board will be deleted. Are you sure
                  want to delete?
                </Label>
              </div>
            </div>
            <div className={styles.apDeletePopupBtnSection}>
              <button
                onClick={(_) => {
                  setApOnDeleteLoader(true);
                  apDeleteItem(apDeletePopup.targetId);
                }}
                className={styles.apDeletePopupYesBtn}
              >
                {apOnDeleteLoader ? <Spinner /> : "Yes"}
              </button>
              <button
                onClick={(_) => {
                  setApDeletePopup({ condition: false, targetId: 0 });
                }}
                className={styles.apDeletePopupNoBtn}
              >
                No
              </button>
            </div>
          </Modal>
        ) : (
          ""
        )}
        {/* Submit-Confirmation-Popup */}
        {submitConfirmationPopup ? (
          <Modal isOpen={submitConfirmationPopup} isBlocking={true}>
            <div
              style={{
                display: "flex",
                justifyContent: "center",
                alignItems: "center",
                marginTop: "30px",
                width: "450px",
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "flex-start",
                  flexDirection: "column",
                  marginBottom: "10px",
                }}
              >
                <Label className={styles.deletePopupTitle}>Confirmation</Label>
                <Label className={styles.deletePopupDesc}>
                  Are you sure want to mark this deliverable as completed?
                </Label>
              </div>
            </div>
            <div className={styles.apDeletePopupBtnSection}>
              <button
                onClick={(_) => {
                  setApSubmitConfirmLoader(true);
                  apUpdateItem(apResponseData.ID);
                }}
                className={styles.apDeletePopupYesBtn}
              >
                {apSubmitConfirmLoader ? <Spinner /> : "Yes"}
              </button>
              <button
                onClick={(_) => {
                  setSubmitConfirmationPopup(false);
                }}
                className={styles.apDeletePopupNoBtn}
              >
                No
              </button>
            </div>
          </Modal>
        ) : (
          ""
        )}
      </div>
    </div>
  );
};

export default AnnualPlan;
