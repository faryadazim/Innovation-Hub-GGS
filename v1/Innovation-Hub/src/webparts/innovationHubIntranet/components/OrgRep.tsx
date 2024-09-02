import * as React from "react";
import { useState, useEffect } from "react";
import { Web } from "@pnp/sp/webs";
import * as moment from "moment";
import {
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  Icon,
  Label,
  Dropdown,
  IDropdownStyles,
  TooltipHost,
  TooltipDelay,
  TooltipOverflowMode,
  DirectionalHint,
  IColumn,
  ILabelStyles,
  Persona,
  PersonaPresence,
  PersonaSize,
  Modal,
  TextField,
  NormalPeoplePicker,
  ITextFieldStyles,
  Spinner,
  PrimaryButton,
  Toggle,
} from "@fluentui/react";

import Service from "./Services";

import "../ExternalRef/styleSheets/Styles.css";
import styles from "./InnovationHubIntranet.module.scss";
import Pagination from "office-ui-fabric-react-pagination";
import { mergeStyles, mergeStyleSets } from "@fluentui/react/lib/Styling";
import CustomLoader from "./CustomLoader";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as Excel from "exceljs/dist/exceljs.min.js";
import * as FileSaver from "file-saver";

import alertify from "alertifyjs";
import "alertifyjs/build/css/alertify.css";
let DateListFormat = "DD/MM/YYYY";

interface IProps {
  context: WebPartContext;
  spcontext: any;
  graphContent: any;
  URL: string;
  peopleList: any;
  isAdmin: boolean;
}
interface IFilter {
  BA: string;
  Title: string;
  Frequency: string;
  Provider: string;
}
interface IDropdown {
  key: string;
  text: string;
}
interface IDropdowns {
  BA: IDropdown[];
  Title: IDropdown[];
  Frequency: IDropdown[];
  Provider: IDropdown[];
}
interface IData {
  ID: number;
  BA: string;
  Title: string;
  Frequency: string;
  ApproverDetails: any[];
  AudienceDetails: any[];
  Provider: any[];
  DisplayStatus: string;
  DueDate: any;
  Confidential: boolean;
}
interface INewData {
  ID: number;

  BA: string;
  Title: string;
  Frequency: string;
  ApproverDetails: any[];
  AudienceDetails: any[];
  Provider: number;

  Confidential: boolean;

  BAValidation: boolean;
  TitleValidation: boolean;
  FrequencyValidation: boolean;
  ProviderValidation: boolean;
  AudienceValidation: boolean;
  ApproverValidation: boolean;

  overAllValidation: boolean;
}
interface IHistoryData {
  ID: any;
  Date: Date;
  comments: string;
  Frequency: string;
  documentLink: string;
  provider: any[];
  audience: any[];
  approver: any[];
  Status: string;
  Year: number;
  TimePeriod: string;
  User: any[];
  UserType: string[];
  SubmittedDate: any;
  ActionedDate: any;
}
interface IGroup {
  key: string;
  name: string;
  startIndex: number;
  count: number;
}
interface IUserDetails {
  userID: number;
  userType: any;
}

let sortORData: IData[] = [];
let sortORFilterData: IData[] = [];

let CurrentPage: number = 1;
let totalPageItems: number = 10;

let arrDatas = [];

const OrgAllReports = (props: IProps): JSX.Element => {
  // variable-Declaration Starts
  const sharepointWeb: any = Web(props.URL);
  const allPeoples: any[] = props.peopleList;

  const ORConfigListName = "Organisation reporting configuration list";
  const OrgReportListName = "OrgReporting";

  const currentLoggedUserEmail: string = props.spcontext.pageContext.user.email;
  const currentLoggedUserID: number = props.peopleList.filter((user) => {
    return user.secondaryText == currentLoggedUserEmail;
  })[0].ID;

  const OR_Year: number = moment().year();
  const OR_WeekNumber: string = `Week ${moment().isoWeek()}`;
  const OR_Month: string = `Month ${moment().format("MMMM")}`;
  const OR_Term: string =
    moment().month() + 1 >= 10
      ? "Term 4"
      : moment().month() + 1 >= 7
      ? "Term 3"
      : moment().month() + 1 >= 4
      ? "Term 2"
      : "Term 1";

  const OR_Term_DueDate: Date =
    OR_Term == "Term 1"
      ? new Date(`${OR_Year}/03/31`)
      : OR_Term == "Term 2"
      ? new Date(`${OR_Year}/06/30`)
      : OR_Term == "Term 3"
      ? new Date(`${OR_Year}/09/30`)
      : OR_Term == "Term 4"
      ? new Date(`${OR_Year}/12/31`)
      : null;

  const ORNewData: INewData = {
    ID: null,

    BA: null,
    Title: null,
    Frequency: null,
    ApproverDetails: [],
    AudienceDetails: [],
    Provider: null,

    Confidential: false,

    BAValidation: false,
    TitleValidation: false,
    FrequencyValidation: false,
    ProviderValidation: false,
    AudienceValidation: false,
    ApproverValidation: false,

    overAllValidation: false,
  };
  const ORModalBoxDrpDwnOptns = {
    BA: [],
    Frequency: [],
  };
  const _ORAllReportsColumn: IColumn[] = props.isAdmin
    ? [
        {
          key: "Column1",
          name: "Business area",
          fieldName: "BA",
          minWidth: 120,
          maxWidth: 120,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) =>
            BAacronymsCollection.filter((ba) => {
              return ba.Name == item.BA;
            })[0].ShortName,
        },
        {
          key: "Column2",
          name: "Title",
          fieldName: "Title",
          minWidth: 100,
          maxWidth: 300,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <div>
              <TooltipHost
                id={item.ID}
                content={item.Title}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.Title}</span>
              </TooltipHost>{" "}
              {item.Confidential ? (
                <Icon
                  iconName="Lock"
                  title="Confidential"
                  className={ORiconStyleClass.infoIcon}
                />
              ) : null}
            </div>
          ),
        },
        {
          key: "Column3",
          name: "Frequency",
          fieldName: "Frequency",
          minWidth: 100,
          maxWidth: 150,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
        },
        {
          key: "Column4",
          name: "Due date",
          fieldName: "DueDate",
          minWidth: 125,
          maxWidth: 170,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => {
            let frequencyType =
              item.Frequency == "Weekly"
                ? "LDW"
                : item.Frequency == "Monthly"
                ? "LDM"
                : item.Frequency == "Term"
                ? "LDT"
                : "";
            let typeAbbreviations =
              item.Frequency == "Weekly"
                ? "Last day of week"
                : item.Frequency == "Monthly"
                ? "Last day of month"
                : item.Frequency == "Term"
                ? "Last day of term"
                : "";
            return (
              <>
                <div
                  style={{
                    marginTop: "1px",
                    cursor: "pointer",
                  }}
                  title={
                    moment(item.DueDate, DateListFormat).format(
                      DateListFormat
                    ) + ` ( ${typeAbbreviations} )`
                  }
                >
                  {moment(item.DueDate, DateListFormat).format(DateListFormat) +
                    ` ( ${frequencyType} )`}
                </div>
              </>
            );
          },
        },
        {
          key: "Column5",
          name: "Provider",
          fieldName: "Provider",
          minWidth: 100,
          maxWidth: 100,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <>
              {item.Provider.length > 0 ? (
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
                      <div
                        title={item.Provider[0].text}
                        style={{ display: "flex" }}
                      >
                        <Persona
                          showOverflowTooltip
                          size={PersonaSize.size24}
                          presence={PersonaPresence.none}
                          showInitialsUntilImageLoads={true}
                          imageUrl={
                            "/_layouts/15/userphoto.aspx?size=S&username=" +
                            `${item.Provider[0].secondaryText}`
                          }
                          styles={{
                            root: {
                              marginLeft: "13px",
                            },
                          }}
                        />
                        <div>
                          {props.isAdmin ? null : (
                            <Label style={{ fontSize: "13px" }}>
                              {item.Provider[0].text}
                            </Label>
                          )}
                        </div>
                        {/* <Label style={{ marginLeft: 10 }}>
                      {item.Provider[0].text}
                    </Label> */}
                      </div>
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
          key: "Column6",
          name: "Audience",
          fieldName: "Audience",
          minWidth: 100,
          maxWidth: 100,
          onRender: (item) => (
            <>
              {item.AudienceDetails.length > 0 ? (
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
                      <div title={item.AudienceDetails[0].text}>
                        <Persona
                          showOverflowTooltip
                          size={PersonaSize.size24}
                          presence={PersonaPresence.none}
                          showInitialsUntilImageLoads={true}
                          imageUrl={
                            "/_layouts/15/userphoto.aspx?size=S&username=" +
                            `${item.AudienceDetails[0].secondaryText}`
                          }
                          styles={{
                            root: {
                              marginRight: "1px",
                            },
                          }}
                        />
                      </div>
                      {item.AudienceDetails.length > 1 ? (
                        <TooltipHost
                          content={
                            <ul style={{ margin: 10, padding: 0 }}>
                              {item.AudienceDetails.map((data, length) => {
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
                                          `${data.secondaryText}`
                                        }
                                      />
                                      <Label style={{ marginLeft: 10 }}>
                                        {data.text}
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
                            {item.AudienceDetails.length}
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
          key: "Column7",
          name: "Approver",
          fieldName: "Approver",
          minWidth: 100,
          maxWidth: 100,
          onRender: (item) => (
            <>
              {item.ApproverDetails.length > 0 ? (
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
                      <div title={item.ApproverDetails[0].text}>
                        <Persona
                          showOverflowTooltip
                          size={PersonaSize.size24}
                          presence={PersonaPresence.none}
                          showInitialsUntilImageLoads={true}
                          imageUrl={
                            "/_layouts/15/userphoto.aspx?size=S&username=" +
                            `${item.ApproverDetails[0].secondaryText}`
                          }
                          styles={{
                            root: {
                              marginRight: "1px",
                            },
                          }}
                        />
                      </div>
                      {item.ApproverDetails.length > 1 ? (
                        <TooltipHost
                          content={
                            <ul style={{ margin: 10, padding: 0 }}>
                              {item.ApproverDetails.map((data, length) => {
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
                                          `${data.secondaryText}`
                                        }
                                      />
                                      <Label style={{ marginLeft: 10 }}>
                                        {data.text}
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
                            {item.ApproverDetails.length}
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
          key: "Column8",
          name: "Status",
          fieldName: "Status",
          minWidth: 200,
          maxWidth: 250,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item: IData) => (
            <>
              {item.DisplayStatus == "Read" ? (
                <div className={ORStatusStyleClass.completed}>
                  {item.DisplayStatus}
                </div>
              ) : item.DisplayStatus == "Scheduled" ? (
                <div className={ORStatusStyleClass.scheduled}>
                  {item.DisplayStatus}
                </div>
              ) : item.DisplayStatus == "Submitted" ? (
                <div className={ORStatusStyleClass.submitted}>
                  {item.DisplayStatus}
                </div>
              ) : item.DisplayStatus == "Over Due" ? (
                <div className={ORStatusStyleClass.overdue}>
                  {item.DisplayStatus}
                </div>
              ) : (
                item.DisplayStatus
              )}
            </>
          ),
        },
       
      ]
    : [
        {
          key: "Column1",
          name: "Business area",
          fieldName: "BA",
          minWidth: 120,
          maxWidth: 120,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) =>
            BAacronymsCollection.filter((ba) => {
              return ba.Name == item.BA;
            })[0].ShortName,
        },
        {
          key: "Column2",
          name: "Title",
          fieldName: "Title",
          minWidth: 100,
          maxWidth: 300,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <div>
              <TooltipHost
                id={item.ID}
                content={item.Title}
                overflowMode={TooltipOverflowMode.Parent}
              >
                <span aria-describedby={item.ID}>{item.Title}</span>
              </TooltipHost>{" "}
              {item.Confidential ? (
                <Icon
                  iconName="Lock"
                  title="Confidential"
                  className={ORiconStyleClass.infoIcon}
                />
              ) : null}
            </div>
          ),
        },
        {
          key: "Column3",
          name: "Frequency",
          fieldName: "Frequency",
          minWidth: 100,
          maxWidth: 150,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
        },
        {
          key: "Column4",
          name: "Due date",
          fieldName: "DueDate",
          minWidth: 125,
          maxWidth: 170,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => {
            let frequencyType =
              item.Frequency == "Weekly"
                ? "LDW"
                : item.Frequency == "Monthly"
                ? "LDM"
                : item.Frequency == "Term"
                ? "LDT"
                : "";
            let typeAbbreviations =
              item.Frequency == "Weekly"
                ? "Last day of week"
                : item.Frequency == "Monthly"
                ? "Last day of month"
                : item.Frequency == "Term"
                ? "Last day of term"
                : "";
            return (
              <>
                <div
                  style={{
                    marginTop: "1px",
                    cursor: "pointer",
                  }}
                  title={
                    moment(item.DueDate, DateListFormat).format(
                      DateListFormat
                    ) + ` ( ${typeAbbreviations} )`
                  }
                >
                  {moment(item.DueDate, DateListFormat).format(DateListFormat) +
                    ` ( ${frequencyType} )`}
                </div>
              </>
            );
          },
        },
        {
          key: "Column5",
          name: "Provider",
          fieldName: "Provider",
          minWidth: 150,
          maxWidth: 250,
          onColumnClick: (
            ev: React.MouseEvent<HTMLElement>,
            column: IColumn
          ) => {
            _onColumnClick(ev, column);
          },
          onRender: (item) => (
            <>
              {item.Provider.length > 0 ? (
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
                      <div
                        title={item.Provider[0].text}
                        style={{ display: "flex" }}
                      >
                        <Persona
                          showOverflowTooltip
                          size={PersonaSize.size24}
                          presence={PersonaPresence.none}
                          showInitialsUntilImageLoads={true}
                          imageUrl={
                            "/_layouts/15/userphoto.aspx?size=S&username=" +
                            `${item.Provider[0].secondaryText}`
                          }
                          styles={{
                            root: {
                              marginLeft: "13px",
                            },
                          }}
                        />
                        <div>
                          <span style={{ fontSize: "13px" }}>
                            {item.Provider[0].text}
                          </span>
                        </div>
                        {/* <Label style={{ marginLeft: 10 }}>
                      {item.Provider[0].text}
                    </Label> */}
                      </div>
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
          key: "Column6",
          name: "Audience",
          fieldName: "Audience",
          minWidth: 100,
          maxWidth: 100,
          onRender: (item) => (
            <>
              {item.AudienceDetails.length > 0 ? (
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
                      <div title={item.AudienceDetails[0].text}>
                        <Persona
                          showOverflowTooltip
                          size={PersonaSize.size24}
                          presence={PersonaPresence.none}
                          showInitialsUntilImageLoads={true}
                          imageUrl={
                            "/_layouts/15/userphoto.aspx?size=S&username=" +
                            `${item.AudienceDetails[0].secondaryText}`
                          }
                          styles={{
                            root: {
                              marginRight: "1px",
                            },
                          }}
                        />
                      </div>
                      {item.AudienceDetails.length > 1 ? (
                        <TooltipHost
                          content={
                            <ul style={{ margin: 10, padding: 0 }}>
                              {item.AudienceDetails.map((data, length) => {
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
                                          `${data.secondaryText}`
                                        }
                                      />
                                      <Label style={{ marginLeft: 10 }}>
                                        {data.text}
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
                            {item.AudienceDetails.length}
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
          key: "Column7",
          name: "Approver",
          fieldName: "Approver",
          minWidth: 100,
          maxWidth: 100,
          onRender: (item) => (
            <>
              {item.ApproverDetails.length > 0 ? (
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
                      <div title={item.ApproverDetails[0].text}>
                        <Persona
                          showOverflowTooltip
                          size={PersonaSize.size24}
                          presence={PersonaPresence.none}
                          showInitialsUntilImageLoads={true}
                          imageUrl={
                            "/_layouts/15/userphoto.aspx?size=S&username=" +
                            `${item.ApproverDetails[0].secondaryText}`
                          }
                          styles={{
                            root: {
                              marginRight: "1px",
                            },
                          }}
                        />
                      </div>
                      {item.ApproverDetails.length > 1 ? (
                        <TooltipHost
                          content={
                            <ul style={{ margin: 10, padding: 0 }}>
                              {item.ApproverDetails.map((data, length) => {
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
                                          `${data.secondaryText}`
                                        }
                                      />
                                      <Label style={{ marginLeft: 10 }}>
                                        {data.text}
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
                            {item.ApproverDetails.length}
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
          key: "Column8",
          name: "View",
          fieldName: "View",
          minWidth: 75,
          maxWidth: 150,

          onRender: (item) => (
            <div style={{ display: "flex" }}>
              <div
                title="History"
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  flexWrap: "wrap",
                  width: 50,
                }}
              >
                <Icon
                  iconName="DocumentReply"
                  className={ORiconStyleClass.historyIcon}
                  onClick={(): void => {
                    getOrgReportHistoryData(item);
                  }}
                />
              </div>
            </div>
          ),
        },
      ];
  const _ORHistoryColumn: IColumn[] = [
    {
      key: "Column1",
      name: "Submitted on",
      fieldName: "Submitted on",
      minWidth: 100,
      maxWidth: 100,
      onRender: (item) => (
        <>
          {item.SubmittedDate ? (
            moment(item.SubmittedDate, DateListFormat).format("DD-MM-YYYY")
          ) : (
            <Label
              style={{
                fontSize: 12,
                width: "100%",
                textAlign: "center",
                alignItems: "center",
              }}
            >
              N/A
            </Label>
          )}
        </>
      ),
    },
    {
      key: "Column2",
      name: "Comments",
      fieldName: "Comments",
      minWidth: 200,
      maxWidth: 300,

      onRender: (item) => (
        <>
          <TooltipHost
            id={item.ID}
            content={item.comments}
            overflowMode={TooltipOverflowMode.Parent}
          >
            <span aria-describedby={item.ID}>{item.comments}</span>
          </TooltipHost>
        </>
      ),
    },
    {
      key: "Column3",
      name: "Open file",
      fieldName: "Open file",
      minWidth: 60,
      maxWidth: 60,
      onRender: (item) => (
        <div
          style={{ width: "100%", textAlign: "center", alignItems: "center" }}
        >
          <a
            href={item.documentLink ? `${item.documentLink}?web=1` : null}
            data-interception="off"
            target="_blank"
          >
            <Icon
              iconName="NavigateExternalInline"
              title="Open document"
              className={
                item.documentLink
                  ? ORiconStyleClass.fileOpenIcon
                  : ORiconStyleClass.fileOpenDisabledIcon
              }
              onClick={() => {}}
            />
          </a>
        </div>
      ),
    },
    {
      key: "Column4",
      name: "User",
      fieldName: "User",
      minWidth: 200,
      maxWidth: 300,
      onRender: (item: IHistoryData) => (
        <>
          {item.User ? (
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "flex-start",
                cursor: "pointer",
              }}
            >
              <div title={item.User[0].text}>
                <Persona
                  showOverflowTooltip
                  size={PersonaSize.size24}
                  presence={PersonaPresence.none}
                  showInitialsUntilImageLoads={true}
                  imageUrl={
                    "/_layouts/15/userphoto.aspx?size=S&username=" +
                    `${item.User[0].secondaryText}`
                  }
                />
              </div>
              <div>
                <Label>{item.User[0].text}</Label>
                {/* {item.UserType.length > 0 ? (
                  <>
                    <TooltipHost
                      content={
                        <ul style={{ margin: 10, padding: 0 }}>
                          {item.UserType.map((type) => {
                            return (
                              <li>
                                <div style={{ display: "flex" }}>
                                  <Label>{type}</Label>
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
                      <Label style={{ cursor: "pointer" }}>
                        {item.User[0].text}
                      </Label>
                    </TooltipHost>
                  </>
                ) : (
                  <Label>{item.User[0].text}</Label>
                )} */}
              </div>
            </div>
          ) : (
            ""
          )}
        </>
      ),
    },
    {
      key: "Column5",
      name: "Type",
      fieldName: "Type",
      minWidth: 200,
      maxWidth: 300,
      onRender: (item) => (
        <>
          <div style={{ display: "flex", flexWrap: "wrap" }}>
            {item.UserType.map((type) => {
              return (
                <div
                  style={{
                    marginRight: 5,
                    marginBottom: 5,
                  }}
                >
                  <Label
                    className={
                      type == "Provider"
                        ? ORHistoryUserTypeStyleClass.Provider
                        : type == "Audience"
                        ? ORHistoryUserTypeStyleClass.Audience
                        : type == "Approver"
                        ? ORHistoryUserTypeStyleClass.Approver
                        : null
                    }
                    // style={{
                    //   border: "1px solid black",
                    //   borderRadius: 20,
                    //   fontSize: 12,
                    //   padding: 5,
                    //   backgroundColor: "#ababab",
                    // }}
                  >
                    {type}
                  </Label>
                </div>
              );
            })}
          </div>
        </>
      ),
    },
    {
      key: "Column6",
      name: "Action",
      fieldName: "Action",
      minWidth: 100,
      maxWidth: 200,
      onRender: (item) => <>{item.Status}</>,
    },
    {
      key: "Column7",
      name: "Actioned on",
      fieldName: "Actioned on",
      minWidth: 100,
      maxWidth: 100,
      onRender: (item) => (
        <>
          {item.ActionedDate ? (
            moment(item.ActionedDate, DateListFormat).format("DD-MM-YYYY")
          ) : (
            <Label
              style={{
                fontSize: 12,
                width: "100%",
                textAlign: "center",
                alignItems: "center",
              }}
            >
              N/A
            </Label>
          )}
        </>
      ),
    },
  ];
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
  const ORFilterKeys: IFilter = {
    BA: "All",
    Title: "All",
    Frequency: "All",
    Provider: "All",
  };
  const ORFilterOptns: IDropdowns = {
    BA: [{ key: "All", text: "All" }],
    Title: [{ key: "All", text: "All" }],
    Frequency: [{ key: "All", text: "All" }],
    Provider: [{ key: "All", text: "All" }],
  };
  // variable-Declaration Ends
  // Style-Section Starts
  const ORfilterLabelStyles: Partial<ILabelStyles> = {
    root: {
      width: 150,
      marginRight: 10,
      fontSize: "13px",
      color: "#323130",
    },
  };
  const ORDropdownStyles: Partial<IDropdownStyles> = {
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
    caretDown: { fontSize: 14, color: "#000" },
  };
  const ORActiveDropdownStyles: Partial<IDropdownStyles> = {
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
    caretDown: { fontSize: 14, color: "#000" },
  };
  const ORModalBoxDropDownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 300,
      margin: "10px 20px",
      backgroundColor: "#fff",
      borderRadius: 4,
    },
    title: {
      backgroundColor: "#fff",
      fontSize: 12,
      borderRadius: 4,
      border: "1px solid #000",
      color: "#000",
    },
    dropdownItemSelected: { fontSize: 12, backgroundColor: "#fff" },
    caretDown: {
      fontSize: 14,
      color: "#000",
    },
    callout: { height: 200 },
  };
  const ORModalBoxReadOnlyDropDownStyles: Partial<IDropdownStyles> = {
    label: {
      color: "#000",
    },
    root: {
      width: 300,
      margin: "10px 20px",
      backgroundColor: "#fff",
      borderRadius: 4,
    },
    title: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
      color: "#000",
      // color: "#7C7C7C",
      border: "1px solid #7C7C7C",
      borderRadius: 4,
    },
    dropdownItemSelected: {
      backgroundColor: "#F5F5F7",
      fontSize: 12,
    },
    caretDown: {
      fontSize: 14,
      color: "#7C7C7C",
      display: "none",
    },
    callout: { height: 200 },
  };
  const ORModalBoxErrorDropDownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 300,
      margin: "10px 20px",
      backgroundColor: "#fff",
      borderRadius: 4,
    },
    title: {
      backgroundColor: "#fff",
      fontSize: 12,
      borderRadius: 4,
      border: "2px solid #f00",
      color: "#000",
    },
    dropdownItemSelected: { fontSize: 12, backgroundColor: "#fff" },
    caretDown: {
      fontSize: 14,
      color: "#000",
    },
    callout: { height: 200 },
  };
  const ORModalBoxWraningDropDownStyles: Partial<IDropdownStyles> = {
    root: {
      width: 300,
      margin: "10px 20px",
      backgroundColor: "#fff",
      borderRadius: 4,
    },
    title: {
      backgroundColor: "#fff",
      fontSize: 12,
      borderRadius: 4,
      border: "2px solid #ff9100",
      color: "#000",
    },
    dropdownItemSelected: { fontSize: 12, backgroundColor: "#fff" },
    caretDown: {
      fontSize: 14,
      color: "#000",
    },
    callout: { height: 200 },
  };
  const ORModalBoxTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: "300px",
      margin: "10px 20px",
    },
    field: {
      fontSize: 12,
      color: "#000",
    },
    fieldGroup: {
      border: "1px solid #000",
    },
  };
  const ORModalBoxReadOnlyTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: 300,
      margin: "10px 20px",
    },
    field: {
      fontSize: 12,
      color: "#000",
      backgroundColor: "#F5F5F7",
    },
    fieldGroup: {
      border: "1px solid #7C7C7C",
    },
  };
  const ORModalBoxErrorTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: "300px",
      margin: "10px 20px",
    },
    field: {
      fontSize: 12,
      color: "#000",
    },
    fieldGroup: {
      border: "2px solid #f00",
    },
  };
  const ORModalBoxWarningTxtBoxStyles: Partial<ITextFieldStyles> = {
    root: {
      width: "300px",
      margin: "10px 20px",
    },
    field: {
      fontSize: 12,
      color: "#000",
    },
    fieldGroup: {
      border: "2px solid #ff9100",
    },
  };
  const ORStatusStyle = mergeStyles({
    textAlign: "center",
    borderRadius: "25px",
    width: "160px",
  });
  const ORStatusStyleClass = mergeStyleSets({
    completed: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#187B29",
        backgroundColor: "#D4FFDB",
      },
      ORStatusStyle,
    ],
    scheduled: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#06637E",
        backgroundColor: "#97E9EC",
      },
      ORStatusStyle,
    ],
    submitted: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#895C09 ",
        backgroundColor: "#FFDB99",
      },
      ORStatusStyle,
    ],
    overdue: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#CB1E06",
        backgroundColor: "#FFD3CD",
      },
      ORStatusStyle,
    ],
    pending: [
      {
        fontWeight: "600",
        padding: "3px",
        color: "#773030",
        backgroundColor: "#e6b1b1",
      },
      ORStatusStyle,
    ],
  });
  const ORHistoryUserTypeStyle = mergeStyles({
    textAlign: "center",
    // width: "160px",
    fontWeight: 600,
    borderRadius: 20,
    fontSize: 12,
    padding: 5,
  });
  const ORHistoryUserTypeStyleClass = mergeStyleSets({
    Provider: [
      {
        color: "#187B29",
        backgroundColor: "#D4FFDB",
        border: "1px solid #187B29",
      },
      ORHistoryUserTypeStyle,
    ],
    Approver: [
      {
        color: "#06637E",
        backgroundColor: "#ccfdff",
        border: "1px solid #06637E",
      },
      ORHistoryUserTypeStyle,
    ],
    Audience: [
      {
        color: "#895C09 ",
        backgroundColor: "#ffeece",
        border: "1px solid #895C09",
      },
      ORHistoryUserTypeStyle,
    ],
  });
  const ORHistorylabelStyles = mergeStyleSets({
    heading: {
      color: "#323130",
      fontSize: 18,
      marginLeft: 10,
      fontWeight: 600,
    },
    DescHeadingLeft: { color: "#000", fontSize: 14, fontWeight: 500 },
    DescHeadingRight: {
      color: "#2392B2",
      fontSize: 14,
      fontWeight: 500,
      marginLeft: 5,
    },
  });
  const ORiconStyleClass = mergeStyleSets({
    refresh: {
      color: "white",
      fontSize: "18px",
      height: 22,
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
    export: {
      color: "#038387",
      fontSize: "18px",
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
    },
    historyIcon: {
      color: "#038387",
      fontSize: 20,
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
      fontWeight: 600,
    },
    editIcon: {
      color: "#038387",
      fontSize: 20,
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
      fontWeight: 600,
    },
    deleteIcon: {
      color: "#038387",
      fontSize: 20,
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
      fontWeight: 600,
    },
    historyBackIcon: {
      color: "#000",
      fontSize: "16px",
      height: 20,
      width: 20,
      cursor: "pointer",
      marginRight: 5,
      fontWeight: 600,
      marginTop: 8,
    },
    fileOpenIcon: {
      userSelect: "none",
      color: "#038387",
      fontSize: 22,
      height: 22,
      width: 22,
      cursor: "pointer",
      marginRight: 5,
    },
    fileOpenDisabledIcon: {
      userSelect: "none",
      color: "#ababab",
      fontSize: 22,
      height: 22,
      width: 22,
      cursor: "not-allowed",
      marginRight: 5,
    },
    infoIcon: {
      userSelect: "none",
      color: "#038387",
      fontSize: 14,
      height: 14,
      width: 14,
      cursor: "pointer",
      marginRight: 5,
    },
  });
  // Style-Section Ends
  // State-Declaration Starts
  const [ORReRender, setORReRender] = useState<boolean>(false);
  const [ORHistoryData, setORHistoryData] = useState<IHistoryData[]>([]);
  const [group, setgroup] = useState<IGroup[]>([]);
  const [ORMasterData, setORMasterData] = useState<IData[]>([]);
  const [ORData, setORData] = useState<IData[]>([]);
  const [ORDisplayData, setORDisplayData] = useState<IData[]>([]);
  const [ORFilter, setORFilter] = useState<IFilter>(ORFilterKeys);
  const [ORFilterData, setORFilterData] = useState<IData[]>([]);
  const [ORFilterDrpDown, setORFilterDrpDown] =
    useState<IDropdowns>(ORFilterOptns);
  const [showHistory, setShowHistory] = useState<{
    condition: boolean;
    data: IData;
  }>({ condition: false, data: null });
  const [ORAddConfigModalBox, setORAddConfigModalBox] = useState<{
    type: string;
    visible: boolean;
    value: INewData;
    oldValue: INewData;
  }>({
    type: "",
    visible: false,
    value: ORNewData,
    oldValue: ORNewData,
  });
  const [ORColumns, setORColumns] = useState<IColumn[]>(_ORAllReportsColumn);
  const [ORModalBoxDropDownOptions, setORModalBoxDropDownOptions] = useState(
    ORModalBoxDrpDwnOptns
  );
  const [ORCurrentPage, setORCurrentPage] = useState<number>(CurrentPage);
  const [ORDuplicateReport, setORDuplicateReport] = useState<boolean>(false);
  const [ORDeletePopup, setORDeletePopup] = useState<{
    condition: boolean;
    targetID: number;
  }>({ condition: false, targetID: null });
  const [ORLoader, setORLoader] = useState("noLoader");
  // State-Declaration Ends
  // Function-Declaration Starts

  // common functions
  const generateExcel = (): void => {
    let arrExport = ORFilterData;
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet("My Sheet");
    worksheet.columns = [
      { header: "Business area", key: "BA", width: 25 },
      { header: "Title", key: "Title", width: 25 },
      { header: "Frequency", key: "Frequency", width: 25 },
      { header: "DueDate", key: "DueDate", width: 20 },
      { header: "Provider", key: "Provider", width: 25 },
      { header: "Audience", key: "Audience", width: 25 },
      { header: "Approver", key: "Approver", width: 25 },
      { header: "Status", key: "Status", width: 60 },
    ];
    arrExport.forEach((item) => {
      let Audience = "";
      item.AudienceDetails.length > 0
        ? item.AudienceDetails.forEach((arr) => {
            Audience += arr.text + ";";
          })
        : null;

      let Approver = "";
      item.ApproverDetails.length > 0
        ? item.ApproverDetails.forEach((arr) => {
            Approver += arr.text + ";";
          })
        : null;

      worksheet.addRow({
        BA: item.BA ? item.BA : "",
        Title: item.Title ? item.Title : "",
        Frequency: item.Frequency ? item.Frequency : "",
        Provider: item.Provider.length > 0 ? item.Provider[0].text : "",
        Status: item.DisplayStatus ? item.DisplayStatus : "",
        DueDate: item.DueDate
          ? moment(item.DueDate, DateListFormat).format("DD/MM/YYYY")
          : "",
        Audience: Audience ? Audience : "",
        Approver: Approver ? Approver : "",
      });
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1"].map((key) => {
      worksheet.getCell(key).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "00e8d1" },
      };
    });
    ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1"].map((key) => {
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
          `Organisationreport-Allreport-${new Date().toLocaleString()}.xlsx`
        )
      )
      .catch((err) => console.log("Error writing excel export", err));
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

  const ErrorPopup = () => (
    alertify.set("notifier", "position", "top-right"),
    alertify.error("Something when error, please contact system admin.")
  );
  const ORErrorFunction = (error: any, functionName: string): void => {
    console.log(error, functionName);

    let response = {
      ComponentName: "Ord reporting - all reports",
      FunctionName: functionName,
      ErrorMessage: JSON.stringify(error["message"]),
      Title: currentLoggedUserEmail,
    };

    Service.SPAddItem({ Listname: "Error Log", RequestJSON: response }).then(
      () => {
        setORAddConfigModalBox({
          type: "",
          visible: false,
          value: ORNewData,
          oldValue: ORNewData,
        });
        setShowHistory({ condition: false, data: null });
        setORLoader("noLoader");
        ErrorPopup();
      }
    );
  };
  const groups = (records: IHistoryData[]): void => {
    let newRecords: { TimePeriod: string; indexValue: number }[] = [];
    records.forEach((item: IHistoryData, index: number) => {
      newRecords.push({
        TimePeriod: item.TimePeriod,
        indexValue: index,
      });
    });

    let varGroup: IGroup[] = [];
    let UniqueMonths = newRecords.reduce((item, e1) => {
      var matches = item.filter((e2) => {
        return e1.TimePeriod === e2.TimePeriod;
      });

      if (matches.length == 0) {
        item.push(e1);
      }
      return item;
    }, []);

    UniqueMonths.forEach((ul) => {
      let TimePeriodLength = newRecords.filter((arr) => {
        return arr.TimePeriod == ul.TimePeriod;
      }).length;
      varGroup.push({
        key: ul.TimePeriod,
        name: ul.TimePeriod,
        startIndex: ul.indexValue,
        count: TimePeriodLength,
      });
    });

    if (records[0].Frequency == "Term" || records[0].Frequency == "Weekly") {
      const sortFilterKeys = (a, b) => {
        if (a.name < b.name) {
          return -1;
        }
        if (a.name > b.name) {
          return 1;
        }
        return 0;
      };

      varGroup.sort(sortFilterKeys);
      setgroup([...varGroup]);
    } else {
      const sortByMonth = (arr) => {
        var months = [
          "January",
          "February",
          "March",
          "April",
          "May",
          "June",
          "July",
          "August",
          "September",
          "October",
          "November",
          "December",
        ];
        arr.sort(function (a, b) {
          let _a = a.name.split(" ");
          let _b = b.name.split(" ");
          return months.indexOf(_a[1]) - months.indexOf(_b[1]);
        });
      };

      sortByMonth(varGroup);
      setgroup([...varGroup]);
    }
  };

  // getData function
  const getConfigData = (): void => {
    sharepointWeb.lists
      .getByTitle(ORConfigListName)
      .items.filter("Inactive ne 1")
      .top(5000)
      .orderBy("ID", false)
      .get()
      .then((items) => {
        // getOrgReportData(items);
        getThresholddata(items);
      })
      .catch((error) => {
        ORErrorFunction(error, "getOrgReportConfig");
      });
  };
  const getThresholddata = (OrgReportConfigData: any[]): void => {
    arrDatas = [];
    let filterCondition = `
    <View Scope='RecursiveAll'>
      <Query>
        <OrderBy>
          <FieldRef Name='Modified' Ascending='FALSE'/>
        </OrderBy>
        <Where>
          <And>
            <Eq>
                <FieldRef Name='MasterData' />
                <Value Type='Text'>Yes</Value>
            </Eq>
            <And>
                <Eq>
                  <FieldRef Name='GroupID' />
                  <Value Type='Number'>1</Value>
                </Eq>
                <And>
                  <Eq>
                      <FieldRef Name='Year' />
                      <Value Type='Number'>${OR_Year}</Value>
                  </Eq>
                  <Neq>
                      <FieldRef Name='Inactive' />
                      <Value Type='Boolean'>1</Value>
                  </Neq>
                </And>
            </And>
          </And>
        </Where>
      </Query>
      <ViewFields>
          <FieldRef Name='ID' />
          <FieldRef Name='Title' />
          <FieldRef Name='ConfigID' />
          <FieldRef Name='MasterData' />
          <FieldRef Name='Frequency' />
          <FieldRef Name='TimePeriod' />
          <FieldRef Name='DueDate' />
          <FieldRef Name='Status' />
          <FieldRef Name='DisplayStatus' />
      </ViewFields>
      <RowLimit Paged='TRUE'>5000</RowLimit>
    </View>`;

    sharepointWeb.lists
      .getByTitle(OrgReportListName)
      .renderListDataAsStream({
        ViewXml: filterCondition,
      })
      .then((data) => {
        arrDatas.push(...data.Row);
        if (arrDatas.length < 5000 && data.NextHref) {
          getPagedValues(data.NextHref, filterCondition, OrgReportConfigData);
        } else {
          processOrgReportData(arrDatas, OrgReportConfigData);
        }
      })
      .catch((err) => {
        ORErrorFunction(err, "getOrgReportData-getThresholddata");
      });
  };
  const getPagedValues = (
    data,
    Filtercondition,
    OrgReportConfigData: any[]
  ): void => {
    sharepointWeb.lists
      .getByTitle(OrgReportListName)
      .renderListDataAsStream({
        ViewXml: Filtercondition,
        Paging: data.substring(1),
      })
      .then(function (data) {
        arrDatas.push(...data.Row);
        if (arrDatas.length < 5000 && data.NextHref) {
          getPagedValues(data.NextHref, Filtercondition, OrgReportConfigData);
        } else {
          processOrgReportData(arrDatas, OrgReportConfigData);
        }
      })
      .catch((err) => {
        ORErrorFunction(err, "getOrgReportData-getPagedValues");
      });
  };
  const processOrgReportData = (
    OrgReportData: any[],
    OrgReportConfigData: any[]
  ) => {
    let _ORdata: IData[] = [];
    if (OrgReportConfigData.length > 0) {
      OrgReportConfigData.forEach((item) => {
        if (item.Confidential) {
          if (
            item.ProviderId == currentLoggedUserID ||
            (item.ApproverId
              ? item.ApproverId.some((app) => app == currentLoggedUserID)
              : false) ||
            (item.AudienceId
              ? item.AudienceId.some((audi) => audi == currentLoggedUserID)
              : false) ||
            props.isAdmin
          ) {
            let filteredArr = OrgReportData.filter((_item) => {
              return (
                item.ID == _item.ConfigID &&
                item.Title == _item.Title &&
                _item.MasterData == "Yes" &&
                ((_item.Frequency == "Term" && _item.TimePeriod == OR_Term) ||
                  (_item.Frequency == "Weekly" &&
                    _item.TimePeriod == OR_WeekNumber) ||
                  (_item.Frequency == "Monthly" &&
                    _item.TimePeriod == OR_Month))
              );
            });

            let _ApproverDetails = [];
            let _AudienceDetails = [];
            let _ProviderDetails = [];

            if (item.ApproverId != null && item.ApproverId.length > 0) {
              item.ApproverId.forEach((user) => {
                _ApproverDetails.push(
                  allPeoples.filter((people) => {
                    return people.ID == user;
                  })[0]
                );
              });
            }

            if (item.AudienceId != null && item.AudienceId.length > 0) {
              item.AudienceId.forEach((user) => {
                _AudienceDetails.push(
                  allPeoples.filter((people) => {
                    return people.ID == user;
                  })[0]
                );
              });
            }

            if (item.ProviderId != null && item.ProviderId) {
              _ProviderDetails.push(
                allPeoples.filter((people) => {
                  return people.ID == item.ProviderId;
                })[0]
              );
            }
            if (filteredArr.length > 0) {
              _ORdata.push({
                ID: item.ID,
                BA: item.BA,
                Title: item.Title,
                Frequency: item.Frequency,
                ApproverDetails: [..._ApproverDetails],
                AudienceDetails: [..._AudienceDetails],
                Provider: [..._ProviderDetails],
                DisplayStatus: filteredArr[0].DisplayStatus,
                DueDate: filteredArr[0].DueDate
                  ? moment(filteredArr[0].DueDate, DateListFormat).format(
                      DateListFormat
                    )
                  : null,
                Confidential: item.Confidential,
              });
            }
          }
        } else {
          let filteredArr = OrgReportData.filter((_item) => {
            return (
              item.ID == _item.ConfigID &&
              item.Title == _item.Title &&
              _item.MasterData == "Yes" &&
              ((_item.Frequency == "Term" && _item.TimePeriod == OR_Term) ||
                (_item.Frequency == "Weekly" &&
                  _item.TimePeriod == OR_WeekNumber) ||
                (_item.Frequency == "Monthly" && _item.TimePeriod == OR_Month))
            );
          });

          let _ApproverDetails = [];
          let _AudienceDetails = [];
          let _ProviderDetails = [];

          if (item.ApproverId != null && item.ApproverId.length > 0) {
            item.ApproverId.forEach((user) => {
              _ApproverDetails.push(
                allPeoples.filter((people) => {
                  return people.ID == user;
                })[0]
              );
            });
          }

          if (item.AudienceId != null && item.AudienceId.length > 0) {
            item.AudienceId.forEach((user) => {
              _AudienceDetails.push(
                allPeoples.filter((people) => {
                  return people.ID == user;
                })[0]
              );
            });
          }

          if (item.ProviderId != null && item.ProviderId) {
            _ProviderDetails.push(
              allPeoples.filter((people) => {
                return people.ID == item.ProviderId;
              })[0]
            );
          }
          if (filteredArr.length > 0) {
            _ORdata.push({
              ID: item.ID,
              BA: item.BA,
              Title: item.Title,
              Frequency: item.Frequency,
              ApproverDetails: [..._ApproverDetails],
              AudienceDetails: [..._AudienceDetails],
              Provider: [..._ProviderDetails],
              DisplayStatus: filteredArr[0].DisplayStatus,
              DueDate: filteredArr[0].DueDate
                ? moment(filteredArr[0].DueDate, DateListFormat).format(
                    DateListFormat
                  )
                : null,
              Confidential: item.Confidential,
            });
          }
        }
      });
    }

    setORFilterData([..._ORdata]);
    sortORFilterData = _ORdata;
    setORData([..._ORdata]);
    sortORData = _ORdata;
    setORMasterData([..._ORdata]);
    reloadFilterDropdowns([..._ORdata]);
    setORDisplayData(_ORdata)
   
    setORLoader("noLoader");
  };
  const getOrgReportData = (items: any[]): void => {
    let _ORdata: IData[] = [];

    // let todayDate: number = parseInt(moment().format("YYYYMMDD"));
    sharepointWeb.lists
      .getByTitle(OrgReportListName)
      .items.select("*", "FieldValuesAsText/DueDate")
      .expand("FieldValuesAsText")

      .filter(
        `MasterData eq 'Yes' and GroupID eq 1 and Year eq ${OR_Year} and Inactive ne 1`
      )
      .top(5000)
      .orderBy("Modified", false)
      .get()
      .then((_items) => {
        if (items.length > 0) {
          items.forEach((item) => {
            if (item.Confidential) {
              if (
                item.ProviderId == currentLoggedUserID ||
                (item.ApproverId
                  ? item.ApproverId.some((app) => app == currentLoggedUserID)
                  : false) ||
                (item.AudienceId
                  ? item.AudienceId.some((audi) => audi == currentLoggedUserID)
                  : false) ||
                props.isAdmin
              ) {
                let filteredArr = _items.filter((_item) => {
                  return (
                    item.ID == _item.ConfigID &&
                    item.Title == _item.Title &&
                    _item.MasterData == "Yes" &&
                    ((_item.Frequency == "Term" &&
                      _item.TimePeriod == OR_Term) ||
                      (_item.Frequency == "Weekly" &&
                        _item.TimePeriod == OR_WeekNumber) ||
                      (_item.Frequency == "Monthly" &&
                        _item.TimePeriod == OR_Month))
                  );
                });

                let _ApproverDetails = [];
                let _AudienceDetails = [];
                let _ProviderDetails = [];

                if (item.ApproverId != null && item.ApproverId.length > 0) {
                  item.ApproverId.forEach((user) => {
                    _ApproverDetails.push(
                      allPeoples.filter((people) => {
                        return people.ID == user;
                      })[0]
                    );
                  });
                }

                if (item.AudienceId != null && item.AudienceId.length > 0) {
                  item.AudienceId.forEach((user) => {
                    _AudienceDetails.push(
                      allPeoples.filter((people) => {
                        return people.ID == user;
                      })[0]
                    );
                  });
                }

                if (item.ProviderId != null && item.ProviderId) {
                  _ProviderDetails.push(
                    allPeoples.filter((people) => {
                      return people.ID == item.ProviderId;
                    })[0]
                  );
                }
                if (filteredArr.length > 0) {
                  _ORdata.push({
                    ID: item.ID,
                    BA: item.BA,
                    Title: item.Title,
                    Frequency: item.Frequency,
                    ApproverDetails: [..._ApproverDetails],
                    AudienceDetails: [..._AudienceDetails],
                    Provider: [..._ProviderDetails],
                    DisplayStatus: filteredArr[0].DisplayStatus,
                    DueDate: moment(
                      filteredArr[0]["FieldValuesAsText"].DueDate,
                      DateListFormat
                    ).format(DateListFormat),
                    Confidential: item.Confidential,
                  });
                }
              }
            } else {
              let filteredArr = _items.filter((_item) => {
                return (
                  item.ID == _item.ConfigID &&
                  item.Title == _item.Title &&
                  _item.MasterData == "Yes" &&
                  ((_item.Frequency == "Term" && _item.TimePeriod == OR_Term) ||
                    (_item.Frequency == "Weekly" &&
                      _item.TimePeriod == OR_WeekNumber) ||
                    (_item.Frequency == "Monthly" &&
                      _item.TimePeriod == OR_Month))
                );
              });

              let _ApproverDetails = [];
              let _AudienceDetails = [];
              let _ProviderDetails = [];

              if (item.ApproverId != null && item.ApproverId.length > 0) {
                item.ApproverId.forEach((user) => {
                  _ApproverDetails.push(
                    allPeoples.filter((people) => {
                      return people.ID == user;
                    })[0]
                  );
                });
              }

              if (item.AudienceId != null && item.AudienceId.length > 0) {
                item.AudienceId.forEach((user) => {
                  _AudienceDetails.push(
                    allPeoples.filter((people) => {
                      return people.ID == user;
                    })[0]
                  );
                });
              }

              if (item.ProviderId != null && item.ProviderId) {
                _ProviderDetails.push(
                  allPeoples.filter((people) => {
                    return people.ID == item.ProviderId;
                  })[0]
                );
              }
              if (filteredArr.length > 0) {
                _ORdata.push({
                  ID: item.ID,
                  BA: item.BA,
                  Title: item.Title,
                  Frequency: item.Frequency,
                  ApproverDetails: [..._ApproverDetails],
                  AudienceDetails: [..._AudienceDetails],
                  Provider: [..._ProviderDetails],
                  DisplayStatus: filteredArr[0].DisplayStatus,
                  DueDate: moment(
                    filteredArr[0]["FieldValuesAsText"].DueDate,
                    DateListFormat
                  ).format(DateListFormat),
                  Confidential: item.Confidential,
                });
              }
            }
          });
        }

        setORFilterData([..._ORdata]);
        sortORFilterData = _ORdata;
        setORData([..._ORdata]);
        sortORData = _ORdata;
        setORMasterData([..._ORdata]);
        reloadFilterDropdowns([..._ORdata]);
        setORDisplayData(_ORdata)

        setORLoader("noLoader");
      })
      .catch((err) => {
        ORErrorFunction(err, "getOrgReportData");
      });
  };
  const getOrgReportHistoryData = (_item: IData): void => {
    let _ORHistoryData: IHistoryData[] = [];
    sharepointWeb.lists
      .getByTitle(OrgReportListName)
      .items.select(
        "*",
        "FieldValuesAsText/SubmittedOn",
        "FieldValuesAsText/ActionedOn"
      )
      .expand("FieldValuesAsText")
      .filter(
        `ConfigID eq ${_item.ID} and Year eq ${OR_Year} and Inactive ne 1`
      )
      .top(5000)
      .orderBy("Modified", false)
      .get()
      .then((items) => {
        items.forEach((item) => {
          if (item.Status && item.Frequency == _item.Frequency) {
            let UserDetails: any[] = [];
            if (item.UserId) {
              UserDetails = props.peopleList.filter((people) => {
                return people.ID == item.UserId;
              });
            }

            _ORHistoryData.push({
              ID: item.ID,
              Date: item.Modified,
              comments: item.Comments,
              Frequency: item.Frequency,
              documentLink: item.DocLink ? item.DocLink : null,
              User: UserDetails,
              UserType: item.UserType,
              provider: [..._item.Provider],
              audience: [..._item.AudienceDetails],
              approver: [..._item.ApproverDetails],
              Status: item.Status,
              Year: item.Year,
              TimePeriod: item.TimePeriod,
              SubmittedDate: item.SubmittedOn
                ? moment(
                    item["FieldValuesAsText"].SubmittedOn,
                    DateListFormat
                  ).format(DateListFormat)
                : null,
              ActionedDate: item.ActionedOn
                ? moment(
                    item["FieldValuesAsText"].ActionedOn,
                    DateListFormat
                  ).format(DateListFormat)
                : null,
            });
          }
        });

        _ORHistoryData.length > 0 ? groups(_ORHistoryData) : null;
        setShowHistory({ condition: true, data: _item });
        setORHistoryData([..._ORHistoryData]);
        setORLoader("noLoader");
      })
      .catch((error) => {
        ORErrorFunction(error, "getOrgReportHistoryData");
      });
  };

  const reloadFilterDropdowns = (data: IData[]): void => {
    let tempArrReload = data;

    const sortFilterKeys = (a, b) => {
      if (a.text < b.text) {
        return -1;
      }
      if (a.text > b.text) {
        return 1;
      }
      return 0;
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

    tempArrReload.forEach((item) => {
      if (
        ORFilterOptns.BA.findIndex((BA) => {
          return BA.key == item.BA;
        }) == -1 &&
        item.BA
      ) {
        ORFilterOptns.BA.push({
          key: item.BA,
          text: item.BA,
        });
      }
      if (
        ORFilterOptns.Title.findIndex((Title) => {
          return Title.key == item.Title;
        }) == -1 &&
        item.Title
      ) {
        ORFilterOptns.Title.push({
          key: item.Title,
          text: item.Title,
        });
      }
      if (
        ORFilterOptns.Frequency.findIndex((Frequency) => {
          return Frequency.key == item.Frequency;
        }) == -1 &&
        item.Frequency
      ) {
        ORFilterOptns.Frequency.push({
          key: item.Frequency,
          text: item.Frequency,
        });
      }

      let tempProvider = [];
      if (item.Provider.length > 0) {
        item.Provider.forEach((people) => {
          tempProvider.push(people.text);
        });

        tempProvider.forEach((_people) => {
          if (
            ORFilterOptns.Provider.findIndex((ProviderOptn) => {
              return ProviderOptn.key == _people;
            }) == -1 &&
            _people != null
          ) {
            ORFilterOptns.Provider.push({
              key: _people,
              text: _people,
            });
          }
        });
      }
    });

    if (
      ORFilterOptns.Provider.some((ProviderOptn) => {
        return (
          ProviderOptn.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      })
    ) {
      ORFilterOptns.Provider.shift();
      let loginUserIndex = ORFilterOptns.Provider.findIndex((user) => {
        return (
          user.text.toLowerCase() ==
          props.spcontext.pageContext.user.displayName.toLowerCase()
        );
      });
      let loginUserData = ORFilterOptns.Provider.splice(loginUserIndex, 1);

      ORFilterOptns.Provider.sort(sortFilterKeys);
      ORFilterOptns.Provider.unshift(loginUserData[0]);
      ORFilterOptns.Provider = usersOrderFunction(ORFilterOptns.Provider);
      ORFilterOptns.Provider.unshift({ key: "All", text: "All" });
    } else {
      ORFilterOptns.Provider.shift();
      ORFilterOptns.Provider.sort(sortFilterKeys);
      ORFilterOptns.Provider = usersOrderFunction(ORFilterOptns.Provider);
      ORFilterOptns.Provider.unshift({ key: "All", text: "All" });
    }

    setORFilterDrpDown(ORFilterOptns);
  };

  const getModalBoxOptions = (): void => {
    //Request Choices
    sharepointWeb.lists
      .getByTitle(ORConfigListName)
      .fields.getByInternalNameOrTitle("BA")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              ORModalBoxDrpDwnOptns.BA.findIndex((rpb) => {
                return rpb.key == choice;
              }) == -1
            ) {
              ORModalBoxDrpDwnOptns.BA.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then()
      .catch((err) => {
        ORErrorFunction(err, "getModalBoxOptions-Request");
      });

    //Documenttype Choices
    sharepointWeb.lists
      .getByTitle(ORConfigListName)
      .fields.getByInternalNameOrTitle("Frequency")()
      .then((response) => {
        response["Choices"].forEach((choice) => {
          if (choice != null) {
            if (
              ORModalBoxDrpDwnOptns.Frequency.findIndex((rdt) => {
                return rdt.key == choice;
              }) == -1
            ) {
              ORModalBoxDrpDwnOptns.Frequency.push({
                key: choice,
                text: choice,
              });
            }
          }
        });
      })
      .then()
      .catch((err) => {
        ORErrorFunction(err, "getModalBoxOptions-Documenttype");
      });

    setORModalBoxDropDownOptions(ORModalBoxDrpDwnOptns);
  };

  // onChange function
  const onChangeFilter = (key: string, option: string): void => {
    let tempData: IData[] = ORData;
    let tempFilterKeys = ORFilter;
    tempFilterKeys[key] = option;
    if (tempFilterKeys.BA != "All") {
      tempData = tempData.filter((arr) => {
        return arr.BA == tempFilterKeys.BA;
      });
    }
    if (tempFilterKeys.Title != "All") {
      tempData = tempData.filter((arr) => {
        return arr.Title == tempFilterKeys.Title;
      });
    }
    if (tempFilterKeys.Frequency != "All") {
      tempData = tempData.filter((arr) => {
        return arr.Frequency == tempFilterKeys.Frequency;
      });
    }

    if (tempFilterKeys.Provider != "All") {
      let devArr = [];
      tempData.forEach((arr) => {
        if (arr.Provider.length != 0) {
          if (
            arr.Provider.some(
              (people) => people.text == tempFilterKeys.Provider
            )
          ) {
            devArr.push(arr);
          }
        }
      });
      tempData = [...devArr];
    }

    setORFilterData([...tempData]);
    sortORFilterData = tempData;
    setORFilter({ ...tempFilterKeys });
       setORDisplayData(tempData)

  };
  const ORAddOnchange = (key, value): void => {
    let tempArronchange = { ...ORAddConfigModalBox.value };
    tempArronchange.overAllValidation = false;
    if (key == "BA") {
      tempArronchange.BA = value;
      tempArronchange.BAValidation = false;
    } else if (key == "Title") {
      tempArronchange.Title = value;
      tempArronchange.TitleValidation = false;
    } else if (key == "Frequency") {
      tempArronchange.Frequency = value;
      tempArronchange.FrequencyValidation = false;
    } else if (key == "ApproverDetails") {
      tempArronchange.ApproverDetails = value;
      tempArronchange.ApproverValidation = false;
      tempArronchange.AudienceValidation = false;
    } else if (key == "Provider") {
      tempArronchange.Provider = value;
      tempArronchange.ProviderValidation = false;
    } else if (key == "AudienceDetails") {
      tempArronchange.AudienceDetails = value;
      tempArronchange.AudienceValidation = false;
      tempArronchange.ApproverValidation = false;
    } else if (key == "Confidential") {
      tempArronchange.Confidential = value;
    }

    setORAddConfigModalBox({
      type: ORAddConfigModalBox.type,
      visible: ORAddConfigModalBox.visible,
      value: { ...tempArronchange },
      oldValue: ORAddConfigModalBox.oldValue,
    });
    setORDuplicateReport(false);
  };

  // data processing function
  const ORValidationFunction = (): void => {
    let tempORAddConfigModalBox = { ...ORAddConfigModalBox };
    if (!tempORAddConfigModalBox.value.BA) {
      tempORAddConfigModalBox.value.BAValidation = true;
      tempORAddConfigModalBox.value.overAllValidation = true;
    }
    if (!tempORAddConfigModalBox.value.Title.trim()) {
      tempORAddConfigModalBox.value.TitleValidation = true;
      tempORAddConfigModalBox.value.overAllValidation = true;
    }
    if (!tempORAddConfigModalBox.value.Frequency) {
      tempORAddConfigModalBox.value.FrequencyValidation = true;
      tempORAddConfigModalBox.value.overAllValidation = true;
    }
    if (
      tempORAddConfigModalBox.value.ApproverDetails.length <= 0 &&
      tempORAddConfigModalBox.value.AudienceDetails.length <= 0
    ) {
      tempORAddConfigModalBox.value.ApproverValidation = true;
      tempORAddConfigModalBox.value.AudienceValidation = true;
      tempORAddConfigModalBox.value.overAllValidation = true;
    }
    if (!tempORAddConfigModalBox.value.Provider) {
      tempORAddConfigModalBox.value.ProviderValidation = true;
      tempORAddConfigModalBox.value.overAllValidation = true;
    }

    if (tempORAddConfigModalBox.value.overAllValidation) {
      setORLoader("noLoader");
      setORAddConfigModalBox({ ...tempORAddConfigModalBox });
    } else {
      findDuplicateReportFunction();
    }
  };
  const findDuplicateReportFunction = (): void => {
    if (
      ORAddConfigModalBox.value.Title &&
      ORAddConfigModalBox.value.Frequency &&
      ORAddConfigModalBox.value.Provider
    ) {
      let filteredArr = ORMasterData.filter((_item: IData) => {
        let _itemProvider: number =
          _item.Provider.length > 0 ? _item.Provider[0].ID : null;
        let AddConfigProvider: number = ORAddConfigModalBox.value.Provider;
        return (
          (_item.Title == ORAddConfigModalBox.value.Title.trim() ||
            _item.Title.toLowerCase() ==
              ORAddConfigModalBox.value.Title.toLowerCase().trim() ||
            _item.Title.toUpperCase() ==
              ORAddConfigModalBox.value.Title.toUpperCase().trim()) &&
          _item.Frequency == ORAddConfigModalBox.value.Frequency &&
          _itemProvider == AddConfigProvider
        );
      });

      if (
        filteredArr.length == 0 ||
        (filteredArr.length == 1 &&
          ORAddConfigModalBox.type == "edit" &&
          filteredArr[0].ID == ORAddConfigModalBox.value.ID)
      ) {
        let _ApproverDetails = [];
        let _AudienceDetails = [];
        if (ORAddConfigModalBox.value.ApproverDetails.length > 0) {
          ORAddConfigModalBox.value.ApproverDetails.forEach((_data) => {
            _ApproverDetails.push(_data.ID);
          });
        }
        if (ORAddConfigModalBox.value.AudienceDetails.length > 0) {
          ORAddConfigModalBox.value.AudienceDetails.forEach((_data) => {
            _AudienceDetails.push(_data.ID);
          });
        }
        let responseData = {
          BA: ORAddConfigModalBox.value.BA,
          Title: ORAddConfigModalBox.value.Title,
          Frequency: ORAddConfigModalBox.value.Frequency,
          ApproverId:
            ORAddConfigModalBox.value.ApproverDetails.length > 0
              ? { results: [..._ApproverDetails] }
              : { results: [] },
          AudienceId:
            ORAddConfigModalBox.value.AudienceDetails.length > 0
              ? { results: [..._AudienceDetails] }
              : { results: [] },
          ProviderId: ORAddConfigModalBox.value.Provider
            ? ORAddConfigModalBox.value.Provider
            : null,
          Confidential: ORAddConfigModalBox.value.Confidential,
        };
        ORAddConfigModalBox.type == "add"
          ? ORAddConfigFunction(responseData)
          : ORUpdateConfigFunction(responseData);
      } else {
        setORDuplicateReport(true);
        setORLoader("noLoader");
      }
    }
  };
  const findUniqueUserDetails = (responseData): IUserDetails[] => {
    let _userDetails: IUserDetails[] = [];
    let _finalUserDetails: IUserDetails[] = [];
    _userDetails.push({
      userID: responseData.ProviderId,
      userType: "Provider",
    });
    responseData.ApproverId.results.length > 0
      ? responseData.ApproverId.results.forEach((approver: number) => {
          _userDetails.push({
            userID: approver,
            userType: "Approver",
          });
        })
      : null;

    responseData.AudienceId.results.length > 0
      ? responseData.AudienceId.results.forEach((audience: number) => {
          _userDetails.push({
            userID: audience,
            userType: "Audience",
          });
        })
      : null;

    _userDetails.forEach((item: IUserDetails) => {
      if (!_finalUserDetails.some((user) => user.userID == item.userID)) {
        let filteredUserDetails: IUserDetails[] = _userDetails.filter(
          (_item) => {
            return _item.userID == item.userID;
          }
        );
        if (filteredUserDetails.length == 1) {
          _finalUserDetails.push({
            userID: item.userID,
            userType: [item.userType],
          });
        } else {
          let newUserType: string[] = [];
          filteredUserDetails.forEach((_item) => {
            newUserType.push(_item.userType);
          });
          newUserType = newUserType.filter((type_, index) => {
            return newUserType.indexOf(type_) === index;
          });
          _finalUserDetails.push({
            userID: item.userID,
            userType: newUserType,
          });
        }
      }
    });

    return _finalUserDetails;
  };

  const ORAddConfigFunction = (responseData): void => {
    sharepointWeb.lists
      .getByTitle(ORConfigListName)
      .items.add(responseData)
      .then((e: any) => {
        ORAddOrgReportFunction(responseData, e.data.ID, "add");
      })
      .catch((err) => {
        ORErrorFunction(err, "ORAddConfigFunction");
      });
  };
  const ORAddOrgReportFunction = (
    responseData,
    configID,
    type: string
  ): void => {
    let _finalUserDetails: IUserDetails[] = findUniqueUserDetails(responseData);

    let processCount = 0;
    let providerFilter = _finalUserDetails.filter((i) => {
      return i.userType.some((type) => type == "Provider") == true;
    });

    // let formatedTermDueDate =
    //   moment(OR_Term_DueDate).day() == 0
    //     ? moment(OR_Term_DueDate).subtract(2, "d").format("YYYY-MM-DD")
    //     : moment(OR_Term_DueDate).day() == 6
    //     ? moment(OR_Term_DueDate).subtract(1, "d").format("YYYY-MM-DD")
    //     : moment(OR_Term_DueDate).format("YYYY-MM-DD");

    let _responseData = {
      ConfigID: configID,
      BA: responseData.BA,
      Title: responseData.Title,
      Frequency: responseData.Frequency,
      DueDate:
        responseData.Frequency == "Weekly"
          ? moment().day(6).format("YYYY-MM-DD")
          : responseData.Frequency == "Monthly"
          ? moment().endOf("month").format("YYYY-MM-DD")
          : responseData.Frequency == "Term"
          ? moment(OR_Term_DueDate).format("YYYY-MM-DD")
          : null,
      ProviderId: providerFilter[0].userID,
      UserId: providerFilter[0].userID,
      UserType:
        providerFilter[0].userType.length > 0
          ? { results: [...providerFilter[0].userType] }
          : { results: [] },
      Status: providerFilter[0].userType.length > 1 ? "Not actioned" : "",
      Year: OR_Year,
      TimePeriod:
        responseData.Frequency == "Term"
          ? OR_Term
          : responseData.Frequency == "Monthly"
          ? OR_Month
          : responseData.Frequency == "Weekly"
          ? OR_WeekNumber.toString()
          : null,

      MasterData: "Yes",
      DisplayStatus: "Scheduled",
      GroupID: 1,
      ApproverCount:
        _finalUserDetails.length == providerFilter.length
          ? 1
          : providerFilter[0].userType.length > 1
          ? _finalUserDetails.length
          : _finalUserDetails.length - 1,
      ApproverActioned: 0,
    };

    sharepointWeb.lists
      .getByTitle(OrgReportListName)
      .items.add(_responseData)
      .then((e) => {
        processCount++;
        if (_finalUserDetails.length != providerFilter.length) {
          _finalUserDetails.forEach((_item) => {
            if (!_item.userType.some((type) => type == "Provider")) {
              let _responseData = {
                ConfigID: configID,
                BA: responseData.BA,
                Title: responseData.Title,
                Frequency: responseData.Frequency,
                DueDate:
                  responseData.Frequency == "Weekly"
                    ? moment().day(6).format("YYYY-MM-DD")
                    : responseData.Frequency == "Monthly"
                    ? moment().endOf("month").format("YYYY-MM-DD")
                    : responseData.Frequency == "Term"
                    ? moment(OR_Term_DueDate).format("YYYY-MM-DD")
                    : null,
                ProviderId: providerFilter[0].userID,
                UserId: _item.userID,
                UserType:
                  _item.userType.length > 0
                    ? { results: [..._item.userType] }
                    : { results: [] },
                Status: "Not actioned",
                Year: OR_Year,
                TimePeriod:
                  responseData.Frequency == "Term"
                    ? OR_Term
                    : responseData.Frequency == "Monthly"
                    ? OR_Month
                    : responseData.Frequency == "Weekly"
                    ? OR_WeekNumber.toString()
                    : null,
                GroupID: 1,
                ParentID: e.data.ID,
              };

              sharepointWeb.lists
                .getByTitle(OrgReportListName)
                .items.add(_responseData)
                .then(() => {
                  processCount++;
                  if (processCount == _finalUserDetails.length) {
                    setORLoader("noLoader");
                    setORAddConfigModalBox({
                      type: "",
                      visible: false,
                      value: ORNewData,
                      oldValue: ORNewData,
                    });
                    setORReRender(!ORReRender);
                  }
                })
                .catch((err) => {
                  ORErrorFunction(err, "ORAddNonProviderData");
                });
            }
          });
        } else if (_finalUserDetails.length == providerFilter.length) {
          setORLoader("noLoader");
          setORAddConfigModalBox({
            type: "",
            visible: false,
            value: ORNewData,
            oldValue: ORNewData,
          });
          setORReRender(!ORReRender);
        }
      })
      .catch((err) => {
        ORErrorFunction(err, "ORAddProviderData");
      });
  };
  const ORAddOrgReportDataMigrateFunction = (
    orgReportData,
    responseData,
    configID: number,
    type: string
  ): void => {
    let finalArr = [];
    let DocLinkArr = orgReportData.filter((report) => {
      return report.DocLink;
    });

    let _DocLink: string = DocLinkArr.length > 0 ? DocLinkArr[0].DocLink : "";

    let SubmittedOnArr = orgReportData.filter((report) => {
      return report.SubmittedOn;
    });
    let _SubmittedOn =
      SubmittedOnArr.length > 0 ? SubmittedOnArr[0].SubmittedOn : "";

    let actionedCount: number = 0;

    let _finalUserDetails: IUserDetails[] = findUniqueUserDetails(responseData);
    // console.log(_finalUserDetails);

    _finalUserDetails.forEach((user) => {
      let filteredOrgReport = orgReportData.filter((report) => {
        return report.UserId == user.userID;
      });

      if (filteredOrgReport.length > 0) {
        // console.log(filteredOrgReport);

        if (
          user.userType.some((type) => type == "Provider") &&
          filteredOrgReport[0].UserType.some((type) => type == "Provider")
        ) {
          let providerTypeCheck: boolean =
            filteredOrgReport[0].UserType.some((type) => type == "Audience") ==
              user.userType.some((type) => type == "Audience") ||
            filteredOrgReport[0].UserType.some((type) => type == "Approver") ==
              user.userType.some((type) => type == "Approver")
              ? true
              : false;
          finalArr.push({
            ConfigID: configID,
            BA: responseData.BA,
            Title: responseData.Title,
            Frequency: responseData.Frequency,
            DueDate:
              responseData.Frequency == "Weekly"
                ? moment().day(6).format("YYYY-MM-DD")
                : responseData.Frequency == "Monthly"
                ? moment().endOf("month").format("YYYY-MM-DD")
                : responseData.Frequency == "Term"
                ? moment(OR_Term_DueDate).format("YYYY-MM-DD")
                : null,
            ProviderId: responseData.ProviderId,
            UserId: user.userID,
            UserType:
              user.userType.length > 0
                ? { results: [...user.userType] }
                : { results: [] },
            Status:
              user.userType.length > 1
                ? providerTypeCheck
                  ? filteredOrgReport[0].Status
                    ? filteredOrgReport[0].Status
                    : "Not actioned"
                  : "Not actioned"
                : "",
            Comments:
              user.userType.length > 1
                ? providerTypeCheck
                  ? filteredOrgReport[0].Comments
                  : ""
                : "",
            SubmittedOn: _SubmittedOn,
            ActionedOn:
              user.userType.length > 1
                ? providerTypeCheck
                  ? filteredOrgReport[0].ActionedOn
                  : null
                : null,
            Year: OR_Year,
            TimePeriod:
              responseData.Frequency == "Term"
                ? OR_Term
                : responseData.Frequency == "Monthly"
                ? OR_Month
                : responseData.Frequency == "Weekly"
                ? OR_WeekNumber.toString()
                : null,

            MasterData: filteredOrgReport[0].MasterData,
            DisplayStatus: filteredOrgReport[0].DisplayStatus, // need attention
            GroupID: 1,
            ApproverCount:
              _finalUserDetails.length == filteredOrgReport.length
                ? 1
                : filteredOrgReport[0].UserType.length > 1
                ? _finalUserDetails.length
                : _finalUserDetails.length - 1,
            ApproverActioned: 0, // need attention
            DocLink: _DocLink,
          });
        } else if (
          (user.userType.some((type) => type == "Approver") &&
            filteredOrgReport[0].UserType.some((type) => type == "Approver")) ||
          (user.userType.some((type) => type == "Audience") &&
            filteredOrgReport[0].UserType.some((type) => type == "Audience"))
        ) {
          finalArr.push({
            ConfigID: configID,
            BA: responseData.BA,
            Title: responseData.Title,
            Frequency: responseData.Frequency,
            DueDate:
              responseData.Frequency == "Weekly"
                ? moment().day(6).format("YYYY-MM-DD")
                : responseData.Frequency == "Monthly"
                ? moment().endOf("month").format("YYYY-MM-DD")
                : responseData.Frequency == "Term"
                ? moment(OR_Term_DueDate).format("YYYY-MM-DD")
                : null,
            ProviderId: responseData.ProviderId,
            UserId: user.userID,
            UserType:
              user.userType.length > 0
                ? { results: [...user.userType] }
                : { results: [] },
            Status: filteredOrgReport[0].Status,
            Year: OR_Year,
            TimePeriod:
              responseData.Frequency == "Term"
                ? OR_Term
                : responseData.Frequency == "Monthly"
                ? OR_Month
                : responseData.Frequency == "Weekly"
                ? OR_WeekNumber.toString()
                : null,
            GroupID: 1,
            ParentID: "", // need attention
            DocLink: _DocLink,
            Comments: filteredOrgReport[0].Comments,
            SubmittedOn: _SubmittedOn,
            ActionedOn: filteredOrgReport[0].ActionedOn
              ? filteredOrgReport[0].ActionedOn
              : null,
          });
        } else {
          finalArr.push({
            ConfigID: configID,
            BA: responseData.BA,
            Title: responseData.Title,
            Frequency: responseData.Frequency,
            DueDate:
              responseData.Frequency == "Weekly"
                ? moment().day(6).format("YYYY-MM-DD")
                : responseData.Frequency == "Monthly"
                ? moment().endOf("month").format("YYYY-MM-DD")
                : responseData.Frequency == "Term"
                ? moment(OR_Term_DueDate).format("YYYY-MM-DD")
                : null,
            ProviderId: responseData.ProviderId,
            UserId: user.userID,
            UserType:
              user.userType.length > 0
                ? { results: [...user.userType] }
                : { results: [] },
            Status: "Not actioned",
            Year: OR_Year,
            TimePeriod:
              responseData.Frequency == "Term"
                ? OR_Term
                : responseData.Frequency == "Monthly"
                ? OR_Month
                : responseData.Frequency == "Weekly"
                ? OR_WeekNumber.toString()
                : null,
            GroupID: 1,
            ParentID: "", // need attention
            DocLink: _DocLink,
            Comments: "",
            SubmittedOn: _SubmittedOn,
            ActionedOn: filteredOrgReport[0].ActionedOn
              ? filteredOrgReport[0].ActionedOn
              : null,
          });
        }
      } else {
        finalArr.push({
          ConfigID: configID,
          BA: responseData.BA,
          Title: responseData.Title,
          Frequency: responseData.Frequency,
          DueDate:
            responseData.Frequency == "Weekly"
              ? moment().day(6).format("YYYY-MM-DD")
              : responseData.Frequency == "Monthly"
              ? moment().endOf("month").format("YYYY-MM-DD")
              : responseData.Frequency == "Term"
              ? moment(OR_Term_DueDate).format("YYYY-MM-DD")
              : null,
          ProviderId: responseData.ProviderId,
          UserId: user.userID,
          UserType:
            user.userType.length > 0
              ? { results: [...user.userType] }
              : { results: [] },
          Status: "Not actioned",
          Year: OR_Year,
          TimePeriod:
            responseData.Frequency == "Term"
              ? OR_Term
              : responseData.Frequency == "Monthly"
              ? OR_Month
              : responseData.Frequency == "Weekly"
              ? OR_WeekNumber.toString()
              : null,
          GroupID: 1,
          ParentID: "", // need attention
          DocLink: _DocLink,
          Comments: "",
          SubmittedOn: _SubmittedOn,
          ActionedOn: null,
        });
      }
    });

    actionedCount = finalArr.filter((arr) => {
      return arr.Status && arr.Status != "Not actioned";
    }).length;

    let processCount = 0;

    let providerFilter = finalArr.filter((arr_) => {
      return arr_.UserType.results.some((type) => type == "Provider");
    });

    let approverCount: number =
      _finalUserDetails.length == providerFilter.length
        ? 1
        : providerFilter[0].UserType.results.length > 1
        ? _finalUserDetails.length
        : _finalUserDetails.length - 1;

    let _responseData = {
      ConfigID: providerFilter[0].ConfigID,
      BA: providerFilter[0].BA,
      Title: providerFilter[0].Title,
      Frequency: providerFilter[0].Frequency,
      DueDate: providerFilter[0].DueDate,
      ProviderId: providerFilter[0].ProviderId,
      UserId: providerFilter[0].UserId,
      UserType: providerFilter[0].UserType,
      Status: providerFilter[0].Status,
      Year: providerFilter[0].Year,
      TimePeriod: providerFilter[0].TimePeriod,
      Comments: providerFilter[0].Comments,
      DocLink: providerFilter[0].DocLink,

      MasterData: "Yes",
      DisplayStatus: approverCount == actionedCount ? "Read" : "Submitted",
      GroupID: 1,
      ApproverCount: approverCount,

      ApproverActioned: actionedCount,
      SubmittedOn: providerFilter[0].SubmittedOn,
      ActionedOn: providerFilter[0].ActionedOn,
    };

    sharepointWeb.lists
      .getByTitle(OrgReportListName)
      .items.add(_responseData)
      .then((e) => {
        processCount++;
        if (_finalUserDetails.length != providerFilter.length) {
          finalArr.forEach((_arr) => {
            if (!_arr.UserType.results.some((type) => type == "Provider")) {
              let _responseData = {
                ConfigID: _arr.ConfigID,
                BA: _arr.BA,
                Title: _arr.Title,
                Frequency: _arr.Frequency,
                DueDate: _arr.DueDate,
                ProviderId: _arr.ProviderId,
                UserId: _arr.UserId,
                UserType: _arr.UserType,
                Status: _arr.Status,
                Year: _arr.Year,
                TimePeriod: _arr.TimePeriod,
                GroupID: 1,
                ParentID: e.data.ID,
                DocLink: _arr.DocLink,
                Comments: _arr.Comments,
                SubmittedOn: _arr.SubmittedOn,
                ActionedOn: _arr.ActionedOn,
              };
              sharepointWeb.lists
                .getByTitle(OrgReportListName)
                .items.add(_responseData)
                .then(() => {
                  processCount++;
                  if (processCount == finalArr.length) {
                    setORLoader("noLoader");
                    setORAddConfigModalBox({
                      type: "",
                      visible: false,
                      value: ORNewData,
                      oldValue: ORNewData,
                    });
                    setORReRender(!ORReRender);
                  }
                })
                .catch((err) => {
                  ORErrorFunction(err, "ORAddNonProviderData");
                });
            }
          });
        } else if (_finalUserDetails.length == providerFilter.length) {
          setORLoader("noLoader");
          setORAddConfigModalBox({
            type: "",
            visible: false,
            value: ORNewData,
            oldValue: ORNewData,
          });
          setORReRender(!ORReRender);
        }
      })
      .catch((err) => {
        ORErrorFunction(
          err,
          "ORAddOrgReportDataMigrateFunction-addProviderData"
        );
      });

    // setORLoader("noLoader");
  };

  const ORUpdateConfigFunction = (responseData): void => {
    let changeLog: string = changeLoggerFunction();
    if (changeLog) {
      sharepointWeb.lists
        .getByTitle(OrgReportListName)
        .items.filter(
          `ConfigID eq ${ORAddConfigModalBox.oldValue.ID} and GroupID eq 1 and Inactive ne 1`
        )
        .top(5000)
        .orderBy("ID", false)
        .get()
        .then((orgReportItems) => {
          let changeLogArr: string[] = changeLog.split("-");
          if (
            changeLogArr.some(
              (log) =>
                log == "Provider" || log == "Audience" || log == "Approver"
            )
          ) {
            // alert("hard changes");
            if (changeLogArr.some((log) => log == "Provider")) {
              // alert("provider changed");
              console.log("provider changed");

              sharepointWeb.lists
                .getByTitle(ORConfigListName)
                .items.getById(ORAddConfigModalBox.oldValue.ID)
                .update(responseData)
                .then(() => {
                  let maxProcessCount: number = 0;
                  orgReportItems.forEach((item) => {
                    sharepointWeb.lists
                      .getByTitle(OrgReportListName)
                      .items.getById(item.ID)
                      .update({
                        Inactive: true,
                        GroupID: 0,
                      })
                      .then((_item) => {
                        maxProcessCount++;
                        if (orgReportItems.length == maxProcessCount) {
                          ORAddOrgReportFunction(
                            responseData,
                            ORAddConfigModalBox.oldValue.ID,
                            "edit"
                          );
                        }
                      })
                      .catch((error) => {
                        ORErrorFunction(
                          error,
                          "ORUpdateConfigFunction-updateOrgReportingData-1"
                        );
                      });
                  });
                })
                .catch((err) => {
                  ORErrorFunction(
                    err,
                    "ORUpdateConfigFunction-updateConfigData-1"
                  );
                });
            } else {
              // alert("provider not changed");
              // console.log("provider not changed");

              sharepointWeb.lists
                .getByTitle(ORConfigListName)
                .items.getById(ORAddConfigModalBox.oldValue.ID)
                .update(responseData)
                .then(() => {
                  let maxProcessCount: number = 0;
                  orgReportItems.forEach((item) => {
                    sharepointWeb.lists
                      .getByTitle(OrgReportListName)
                      .items.getById(item.ID)
                      .update({
                        Inactive: true,
                      })
                      .then((_item) => {
                        maxProcessCount++;
                        if (orgReportItems.length == maxProcessCount) {
                          let providerObj = orgReportItems.filter((_obj) => {
                            return _obj.MasterData == "Yes";
                          })[0];

                          if (providerObj.DisplayStatus == "Scheduled") {
                            ORAddOrgReportFunction(
                              responseData,
                              ORAddConfigModalBox.oldValue.ID,
                              "edit"
                            );
                          } else {
                            // alert("status changed");
                            ORAddOrgReportDataMigrateFunction(
                              orgReportItems,
                              responseData,
                              ORAddConfigModalBox.oldValue.ID,
                              "edit"
                            );
                          }
                        }
                      })
                      .catch((error) => {
                        ORErrorFunction(
                          error,
                          "ORUpdateConfigFunction-updateOrgReportingData-1"
                        );
                      });
                  });
                })
                .catch((err) => {
                  ORErrorFunction(
                    err,
                    "ORUpdateConfigFunction-updateConfigData-1"
                  );
                });
            }
          } else {
            // alert("soft changes");

            sharepointWeb.lists
              .getByTitle(ORConfigListName)
              .items.getById(ORAddConfigModalBox.oldValue.ID)
              .update({
                BA: ORAddConfigModalBox.value.BA,
                Confidential: ORAddConfigModalBox.value.Confidential,
              })
              .then(() => {
                let maxProcessCount: number = 0;
                orgReportItems.forEach((item) => {
                  maxProcessCount++;
                  sharepointWeb.lists
                    .getByTitle(OrgReportListName)
                    .items.getById(item.ID)
                    .update({
                      BA: ORAddConfigModalBox.value.BA,
                    })
                    .then(() => {
                      if (orgReportItems.length == maxProcessCount) {
                        setORReRender(!ORReRender);
                        setORLoader("noLoader");
                        setORAddConfigModalBox({
                          type: "",
                          visible: false,
                          value: ORNewData,
                          oldValue: ORNewData,
                        });
                      }
                    })
                    .catch((error) => {
                      ORErrorFunction(
                        error,
                        "ORUpdateConfigFunction-updateOrgReportingData-3"
                      );
                    });
                });
              })
              .catch((error) => {
                ORErrorFunction(
                  error,
                  "ORUpdateConfigFunction-updateConfigData-3"
                );
              });
          }
        })
        .catch((err) => {
          ORErrorFunction(err, "ORUpdateConfigFunction-getOrgReportingData");
        });
    } else {
      setORLoader("noLoader");
    }
  };
  const ORDeleteFunction = (configID: number): void => {
    sharepointWeb.lists
      .getByTitle(ORConfigListName)
      .items.getById(configID)
      .update({
        Inactive: true,
      })
      .then((_item) => {
        sharepointWeb.lists
          .getByTitle(OrgReportListName)
          .items.filter(
            `ConfigID eq ${configID} and GroupID eq 1 and Inactive ne 1`
          )
          .top(5000)
          .orderBy("ID", false)
          .get()
          .then((orgReportItems) => {
            let maxProcessCount: number = 0;
            orgReportItems.forEach((item) => {
              sharepointWeb.lists
                .getByTitle(OrgReportListName)
                .items.getById(item.ID)
                .update({
                  Inactive: true,
                  GroupID: 0,
                })
                .then((_item) => {
                  maxProcessCount++;
                  if (orgReportItems.length == maxProcessCount) {
                    setORReRender(!ORReRender);
                    setORLoader("noLoader");
                    setORDeletePopup({
                      condition: false,
                      targetID: null,
                    });
                  }
                })
                .catch((error) => {
                  ORErrorFunction(
                    error,
                    "ORDeleteFunction-updateOrgReportItem"
                  );
                });
            });
          })
          .catch((err) => {
            ORErrorFunction(err, "ORDeleteFunction-getOrgReportItems");
          });
      })
      .catch((err) => {
        ORErrorFunction(err, "ORDeleteFunction-updateOrgReportConfigItem");
      });
  };

  const changeLoggerFunction = (): string => {
    let changeLog: string[] = [];

    if (ORAddConfigModalBox.oldValue.BA != ORAddConfigModalBox.value.BA) {
      changeLog.push("BA");
    }
    if (
      ORAddConfigModalBox.oldValue.Provider !=
      ORAddConfigModalBox.value.Provider
    ) {
      changeLog.push("Provider");
    }

    if (
      ORAddConfigModalBox.oldValue.AudienceDetails.length !=
      ORAddConfigModalBox.value.AudienceDetails.length
    ) {
      changeLog.push("Audience");
    } else if (
      ORAddConfigModalBox.oldValue.AudienceDetails.length != 0 &&
      ORAddConfigModalBox.value.AudienceDetails.length != 0
    ) {
      let res: boolean = false;
      for (
        let i = 0;
        i < ORAddConfigModalBox.oldValue.AudienceDetails.length;
        i++
      ) {
        let _arr = ORAddConfigModalBox.oldValue.AudienceDetails[i];
        if (
          !ORAddConfigModalBox.value.AudienceDetails.some(
            (obj) => obj.ID == _arr.ID
          )
        ) {
          res = true;
          changeLog.push("Audience");
          break;
        }
      }
    }

    if (
      ORAddConfigModalBox.oldValue.ApproverDetails.length !=
      ORAddConfigModalBox.value.ApproverDetails.length
    ) {
      changeLog.push("Approver");
    } else if (
      ORAddConfigModalBox.oldValue.ApproverDetails.length != 0 &&
      ORAddConfigModalBox.value.ApproverDetails.length != 0
    ) {
      let res: boolean = false;
      for (
        let i = 0;
        i < ORAddConfigModalBox.oldValue.ApproverDetails.length;
        i++
      ) {
        let _arr = ORAddConfigModalBox.oldValue.ApproverDetails[i];
        if (
          !ORAddConfigModalBox.value.ApproverDetails.some(
            (obj) => obj.ID == _arr.ID
          )
        ) {
          res = true;
          changeLog.push("Approver");
          break;
        }
      }
    }

    if (
      ORAddConfigModalBox.oldValue.Confidential !=
      ORAddConfigModalBox.value.Confidential
    ) {
      changeLog.push("Confidential");
    }

    return changeLog.length > 0 ? changeLog.join("-") : "";
  };
  // column-sorting function
  const _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const tempORColumns = _ORAllReportsColumn;
    const newColumns: IColumn[] = tempORColumns.slice();
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

    const newORData = _copyAndSort(
      sortORData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    const newORFilterData = _copyAndSort(
      sortORFilterData,
      currColumn.fieldName!,
      currColumn.isSortedDescending
    );
    setORData([...newORData]);
    setORFilterData([...newORFilterData]);
    setORDisplayData(newORData)

   
  };
  function _copyAndSort<T>(
    items: T[],
    columnKey: string,
    isSortedDescending?: boolean
  ): T[] {
    let key = columnKey as keyof T;
    return items
      .slice(0)
      .sort((a: T, b: T) =>
        (isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1
      );
  }

  // Function-Declaration Ends

  useEffect(() => {
    setORLoader("StartLoader");
    getModalBoxOptions();
    getConfigData();
  }, [ORReRender]);
  return (
    <>
      {showHistory.condition ? (
        <div>
          <div style={{ padding: 10, marginTop: 10 }}>
            <div style={{ display: "flex", alignItems: "center" }}>
              <Icon
                iconName="ChromeBack"
                className={ORiconStyleClass.historyBackIcon}
                onClick={(): void => {
                  setShowHistory({ condition: false, data: null });
                }}
              />
              <Label className={ORHistorylabelStyles.heading}>History</Label>
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                marginTop: "10px",
                marginBottom: "5px",
                flexWrap: "wrap",
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  marginRight: 25,
                }}
              >
                <Label className={ORHistorylabelStyles.DescHeadingLeft}>
                  Business area :
                </Label>
                <Label className={ORHistorylabelStyles.DescHeadingRight}>
                  {showHistory.data.BA}
                </Label>
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  marginRight: 25,
                }}
              >
                <Label className={ORHistorylabelStyles.DescHeadingLeft}>
                  Title :
                </Label>
                <Label className={ORHistorylabelStyles.DescHeadingRight}>
                  {showHistory.data.Title}
                </Label>
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  marginRight: 25,
                }}
              >
                <Label className={ORHistorylabelStyles.DescHeadingLeft}>
                  Frequency :
                </Label>
                <Label className={ORHistorylabelStyles.DescHeadingRight}>
                  {showHistory.data.Frequency}
                </Label>
              </div>
              {showHistory.data.Provider.length > 0 ? (
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    marginRight: 25,
                  }}
                >
                  <Label className={ORHistorylabelStyles.DescHeadingLeft}>
                    Provider :
                  </Label>
                  <Label className={ORHistorylabelStyles.DescHeadingRight}>
                    {showHistory.data.Provider[0].text}
                  </Label>
                </div>
              ) : null}
            </div>
            <div>
              <DetailsList
                items={ORHistoryData}
                columns={_ORHistoryColumn}
                groups={group}
                groupProps={{
                  showEmptyGroups: true,
                }}
                styles={{
                  root: {
                    ".ms-DetailsHeader-cellTitle": {
                      // justifyContent: "center !important",
                    },
                    ".ms-DetailsRow-cell": {
                      display: "flex",
                      alignItems: "center",
                      // justifyContent: "center",
                    },
                  },
                }}
                setKey="set"
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
              />
            </div>
            {ORHistoryData.length == 0 && (
             
               <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  marginTop: "15px",
                }}
              >
                <Label style={{ color: "#2392B2" }}>No data found !!!</Label>
              </div>
             
             
            )}
          </div>
        </div>
      ) : (
        <div>
          {ORLoader == "StartLoader" ? (
            <CustomLoader />
          ) : (
            <div>
              {/* Header-Section Starts */}
                {/* Header-Btn-Section Starts */}
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "flex-start",
                    // paddingTop: 16,
                    paddingBottom: "10px",
                  }}
                >
                   <Label style={{ color: "#000", fontSize: 24, padding: 0 }}>
          
          </Label>
                  <Label
                    onClick={() => {
                      generateExcel();
                    }}
                    style={{
                      backgroundColor: "#EBEBEB",
                      padding: "7px 15px",
                      cursor: "pointer",
                      fontSize: "12px",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      borderRadius: "3px",
                      color: "#1D6F42",
                      marginRight: 10,
                    }}
                  >
                    <Icon
                      style={{
                        color: "#1D6F42",
                      }}
                      iconName="ExcelDocument"
                      className={ORiconStyleClass.export}
                    />
                    Export as XLS
                  </Label>
                 
                </div>
                {/* Header-Btn-Section Ends */}
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "end",
                  flexWrap: "wrap",
                }}
              >
                {/* Filter-Section Starts */}
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    marginBottom: 10,
                    flexWrap: "wrap",
                  }}
                >
                  <div>
                    <Label styles={ORfilterLabelStyles}>Business area</Label>
                    <Dropdown
                      placeholder="Select an option"
                      options={ORFilterDrpDown.BA}
                      selectedKey={ORFilter.BA}
                      styles={
                        ORFilter.BA == "All"
                          ? ORDropdownStyles
                          : ORActiveDropdownStyles
                      }
                      onChange={(e, option: any) => {
                        onChangeFilter("BA", option["key"]);
                      }}
                    />
                  </div>
                  <div>
                    <Label styles={ORfilterLabelStyles}>Title</Label>
                    <Dropdown
                      placeholder="Select an option"
                      options={ORFilterDrpDown.Title}
                      selectedKey={ORFilter.Title}
                      styles={
                        ORFilter.Title == "All"
                          ? ORDropdownStyles
                          : ORActiveDropdownStyles
                      }
                      onChange={(e, option: any) => {
                        onChangeFilter("Title", option["key"]);
                      }}
                    />
                  </div>
                  <div>
                    <Label styles={ORfilterLabelStyles}>Frequency</Label>
                    <Dropdown
                      placeholder="Select an option"
                      options={ORFilterDrpDown.Frequency}
                      selectedKey={ORFilter.Frequency}
                      styles={
                        ORFilter.Frequency == "All"
                          ? ORDropdownStyles
                          : ORActiveDropdownStyles
                      }
                      onChange={(e, option: any) => {
                        onChangeFilter("Frequency", option["key"]);
                      }}
                    />
                  </div>
                  <div>
                    <Label styles={ORfilterLabelStyles}>Provider</Label>
                    <Dropdown
                      placeholder="Select an option"
                      options={ORFilterDrpDown.Provider}
                      selectedKey={ORFilter.Provider}
                      styles={
                        ORFilter.Provider == "All"
                          ? ORDropdownStyles
                          : ORActiveDropdownStyles
                      }
                      onChange={(e, option: any) => {
                        onChangeFilter("Provider", option["key"]);
                      }}
                    />
                  </div>
                  <div>
                    <Icon
                      iconName="Refresh"
                      title="Click to reset"
                      className={ORiconStyleClass.refresh}
                      onClick={() => {
                        setORFilterData([...ORMasterData]);
                        setORData([...ORMasterData]);
                        sortORFilterData = ORMasterData;
                        sortORData = ORMasterData;
                        setORDisplayData(ORMasterData)
                    
                        setORFilter({ ...ORFilterKeys });
                        setORColumns(_ORAllReportsColumn);
                      }}
                    />
                  </div>
                </div>
                {/* Filter-Section Ends */}
              
              </div>
              {/* Header-Section Ends */}
              {/* Body-Section Starts */}
              <div>
                {/* DetailList-Section Starts */}
                <DetailsList
                  items={ORDisplayData}
                  columns={ORColumns}
                  styles={{
                    root: {
                      ".ms-DetailsRow-cell": {
                        // display: "flex",
                        // alignItems: "center",
                        height: 40,
                      },
                    },
                  }}
                  setKey="set"
                  layoutMode={DetailsListLayoutMode.justified}
                  selectionMode={SelectionMode.none}
                />
                {/* DetailList-Section Ends */}
              </div>
              {ORFilterData.length == 0 && (
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
            {/* Body-Section Ends */}
          </div>
              ) }
          {/* Modal-Section Starts */}
          {ORAddConfigModalBox.visible ? (
            <Modal isOpen={ORAddConfigModalBox.visible} isBlocking={false}>
              <div style={{ padding: "15px 20px" }}>
                <div
                  style={{
                    fontSize: 24,
                    textAlign: "center",
                    color: "#2392B2",
                    fontWeight: "600",
                    marginBottom: "20px",
                  }}
                >
                  {ORAddConfigModalBox.type == "add"
                    ? "Add report"
                    : "Update report"}
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
                      placeholder="Select an option"
                      options={ORModalBoxDropDownOptions.BA}
                      styles={
                        ORAddConfigModalBox.value.BAValidation
                          ? ORModalBoxErrorDropDownStyles
                          : ORModalBoxDropDownStyles
                      }
                      selectedKey={ORAddConfigModalBox.value.BA}
                      onChange={(e, option: any) => {
                        ORAddOnchange("BA", option["key"]);
                      }}
                    />
                  </div>
                  <div>
                    <TextField
                      label="Title"
                      required={
                        ORAddConfigModalBox.type == "edit" ? false : true
                      }
                      readOnly={ORAddConfigModalBox.type == "edit"}
                      placeholder="Add new project"
                      value={ORAddConfigModalBox.value.Title}
                      styles={
                        ORAddConfigModalBox.type == "edit"
                          ? ORModalBoxReadOnlyTxtBoxStyles
                          : ORDuplicateReport
                          ? ORModalBoxWarningTxtBoxStyles
                          : ORAddConfigModalBox.value.TitleValidation
                          ? ORModalBoxErrorTxtBoxStyles
                          : ORModalBoxTxtBoxStyles
                      }
                      className={styles.projectField}
                      onChange={(e, value: string) => {
                        ORAddOnchange("Title", value);
                      }}
                    />
                  </div>
                  <div>
                    <Dropdown
                      label="Frequency"
                      placeholder="Select an option"
                      disabled={ORAddConfigModalBox.type == "edit"}
                      options={ORModalBoxDropDownOptions.Frequency}
                      styles={
                        ORAddConfigModalBox.type == "edit"
                          ? ORModalBoxReadOnlyDropDownStyles
                          : ORDuplicateReport
                          ? ORModalBoxWraningDropDownStyles
                          : ORAddConfigModalBox.value.FrequencyValidation
                          ? ORModalBoxErrorDropDownStyles
                          : ORModalBoxDropDownStyles
                      }
                      selectedKey={ORAddConfigModalBox.value.Frequency}
                      onChange={(e, option: any) => {
                        ORAddOnchange("Frequency", option["key"]);
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
                    <Label
                      style={{
                        transform: "translate(20px, 10px)",
                      }}
                    >
                      Provider
                    </Label>
                    <NormalPeoplePicker
                      className={styles.orgConfigModalPeoplePicker}
                      styles={
                        ORDuplicateReport
                          ? {
                              root: {
                                selectors: {
                                  ".ms-BasePicker-text": {
                                    border: "2px solid #ff9100",
                                  },
                                },
                              },
                            }
                          : ORAddConfigModalBox.value.ProviderValidation
                          ? {
                              root: {
                                selectors: {
                                  ".ms-BasePicker-text": {
                                    border: "2px solid #f00",
                                  },
                                },
                              },
                            }
                          : {
                              root: {
                                selectors: {
                                  selectors: {
                                    ".ms-BasePicker-text": {
                                      border: "1px solid #000",
                                    },
                                  },
                                },
                              },
                            }
                      }
                      onResolveSuggestions={GetUserDetails}
                      itemLimit={1}
                      selectedItems={allPeoples.filter((people) => {
                        return people.ID == ORAddConfigModalBox.value.Provider;
                      })}
                      onChange={(selectedUser) => {
                        selectedUser.length != 0
                          ? ORAddOnchange("Provider", selectedUser[0]["ID"])
                          : ORAddOnchange("Provider", "");
                      }}
                    />
                  </div>
                  <div>
                    <Label
                      style={{
                        transform: "translate(20px, 10px)",
                      }}
                    >
                      Audience
                    </Label>
                    <NormalPeoplePicker
                      className={styles.orgConfigModalPeoplePicker}
                      styles={
                        ORAddConfigModalBox.value.AudienceValidation
                          ? {
                              root: {
                                selectors: {
                                  ".ms-BasePicker-text": {
                                    border: "2px solid #f00",
                                  },
                                },
                              },
                            }
                          : {
                              root: {
                                selectors: {
                                  selectors: {
                                    ".ms-BasePicker-text": {
                                      border: "1px solid #000",
                                    },
                                  },
                                },
                              },
                            }
                      }
                      onResolveSuggestions={GetUserDetails}
                      itemLimit={5}
                      selectedItems={ORAddConfigModalBox.value.AudienceDetails}
                      onChange={(selectedUser) => {
                        selectedUser.length != 0
                          ? ORAddOnchange("AudienceDetails", selectedUser)
                          : ORAddOnchange("AudienceDetails", "");
                      }}
                    />
                  </div>
                  <div>
                    <Label
                      style={{
                        transform: "translate(20px, 10px)",
                      }}
                    >
                      Approver
                    </Label>
                    <NormalPeoplePicker
                      className={styles.orgConfigModalPeoplePicker}
                      styles={
                        ORAddConfigModalBox.value.ApproverValidation
                          ? {
                              root: {
                                selectors: {
                                  ".ms-BasePicker-text": {
                                    border: "2px solid #f00",
                                  },
                                },
                              },
                            }
                          : {
                              root: {
                                selectors: {
                                  selectors: {
                                    ".ms-BasePicker-text": {
                                      border: "1px solid #000",
                                    },
                                  },
                                },
                              },
                            }
                      }
                      onResolveSuggestions={GetUserDetails}
                      itemLimit={5}
                      selectedItems={ORAddConfigModalBox.value.ApproverDetails}
                      onChange={(selectedUser) => {
                        selectedUser.length != 0
                          ? ORAddOnchange("ApproverDetails", selectedUser)
                          : ORAddOnchange("ApproverDetails", "");
                      }}
                    />
                  </div>
                </div>
                <div
                  style={{ display: "flex", justifyContent: "space-between" }}
                >
                  <div
                    style={{
                      marginTop: 30,
                      marginLeft: 20,
                      position: "relative",
                    }}
                  >
                    {/* <Label
                      style={{
                        transform: "translate(20px, 10px)",
                      }}
                    >
                      Sample
                    </Label> */}
                    <Toggle
                      label={
                        <div
                          style={{
                            position: "absolute",
                            left: "0",
                            top: "0",
                            width: "200px",
                          }}
                        >
                          Confidential
                        </div>
                      }
                      inlineLabel
                      style={{ transform: "translateX(100px)" }}
                      checked={ORAddConfigModalBox.value.Confidential}
                      onChange={() => {
                        ORAddOnchange(
                          "Confidential",
                          !ORAddConfigModalBox.value.Confidential
                        );
                      }}
                    />
                  </div>
                  <div className={styles.ORModalBoxButtonSection}>
                    {ORDuplicateReport ? (
                      <Label style={{ color: "#ff9100", fontWeight: 600 }}>
                        * This report already exists
                      </Label>
                    ) : null}
                    {ORAddConfigModalBox.value.overAllValidation ? (
                      <Label style={{ color: "#f00", fontWeight: 600 }}>
                        * All fields are mandatory
                      </Label>
                    ) : null}
                    <button
                      className={styles.ORModalBoxSubmitBtn}
                      onClick={(_) => {
                        if (ORLoader == "noLoader") {
                          setORLoader("onModalSubmit");
                          ORValidationFunction();
                        }
                      }}
                      style={{ display: "flex" }}
                    >
                      {ORLoader == "onModalSubmit" ? (
                        <Spinner />
                      ) : (
                        <span>
                          <Icon
                            iconName="Save"
                            style={{ position: "relative", top: 3, left: -8 }}
                          />
                          {ORAddConfigModalBox.type == "add"
                            ? "Submit"
                            : "Update"}
                        </span>
                      )}
                    </button>
                    <button
                      className={styles.ORModalBoxBackBtn}
                      onClick={(_) => {
                        if (ORLoader == "noLoader") {
                          setORDuplicateReport(false);
                          setORAddConfigModalBox({
                            type: "",
                            visible: false,
                            value: ORNewData,
                            oldValue: ORNewData,
                          });
                        }
                      }}
                    >
                      <span>
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
          ) : null}

          {ORDeletePopup.condition ? (
            <Modal isOpen={ORDeletePopup.condition} isBlocking={false}>
              <div>
                <div
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                    marginTop: 30,
                    width: 450,
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "flex-start",
                      flexDirection: "column",
                      marginBottom: 10,
                    }}
                  >
                    <Label className={styles.ORDeletePopupTitle}>
                      Delete Report
                    </Label>
                    <Label className={styles.ORDeletePopupDesc}>
                      Are you sure want to delete this report?
                    </Label>
                  </div>
                </div>
                <div className={styles.ORDeletePopupBtnSection}>
                  <button
                    onClick={(_) => {
                      if (ORLoader != "deletePopupLoader") {
                        setORLoader("deletePopupLoader");
                        ORDeleteFunction(ORDeletePopup.targetID);
                      }
                    }}
                    className={styles.ORDeletePopupSuccessBtn}
                  >
                    {ORLoader == "deletePopupLoader" ? <Spinner /> : "Yes"}
                  </button>
                  <button
                    onClick={(_) => {
                      ORLoader != "deletePopupLoader"
                        ? setORDeletePopup({ condition: false, targetID: null })
                        : null;
                    }}
                    className={styles.ORDeletePopupCancelBtn}
                  >
                    No
                  </button>
                </div>
              </div>
            </Modal>
          ) : null}
          {/* Modal-Section Ends */}
        </div>
      )}
    </>
  );
};

export default OrgAllReports;
