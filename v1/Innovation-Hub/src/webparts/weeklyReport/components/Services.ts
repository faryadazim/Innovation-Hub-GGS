import * as React from "react";
import { useState, useEffect } from "react";
import { sp } from "@pnp/sp/presets/all";
import {
  IFilter,
  IListItems,
  IAddList,
  IUpdateList,
  IDeleteList,
} from "./IServiceProps";
import { IItemAddResult } from "@pnp/sp/items";
import "@pnp/sp/webs";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";

var webURL;

if (window.location.href.toLowerCase().indexOf("production/") > -1) {
  /* Production URL */
  webURL = "https://ggsaus.sharepoint.com";
} else {
  /* Development URL */
  webURL = "https://ggsaus.sharepoint.com/sites/Intranet_Test";
}

const _webURL = Web(webURL);

const SPAddItem = async (params: IAddList): Promise<IItemAddResult> => {
  return await _webURL.lists
    .getByTitle(params.Listname)
    .items.add(params.RequestJSON);
};
const SPUpdateItem = async (params: IUpdateList): Promise<IItemAddResult> => {
  return await _webURL.lists
    .getByTitle(params.Listname)
    .items.getById(params.ID)
    .update(params.RequestJSON);
};
const SPDeleteItem = async (params: IDeleteList): Promise<void> => {
  return await _webURL.lists
    .getByTitle(params.Listname)
    .items.getById(params.ID)
    .delete();
};
const SPReadItems = async (params: IListItems): Promise<[]> => {
  params = formatInputs(params);
  let filterValue: string = formatFilterValue(params.Filter);

  return await _webURL.lists
    .getByTitle(params.Listname)
    .items.select(params.Select)
    .filter(filterValue)
    .expand(params.Expand)
    .top(params.Topcount)
    .orderBy(params.Orderby, params.Orderbydecorasc)
    .get();
};

const readItemsFromSharepointListForDashbaord = async (
  params: IListItems
): Promise<[]> => {
  params = formatInputs(params);
  let filterValue: string = formatFilterValue(params.Filter);
  let skipcount = params.PageNumber * params.PageCount - params.PageCount;

  return await _webURL.lists
    .getByTitle(params.Listname)
    .items.select(params.Select)
    .filter(filterValue)
    .expand(params.Expand)
    .skip(skipcount)
    .top(params.PageCount)
    .orderBy(params.Orderby, params.Orderbydecorasc)
    .get();
};

const formatInputs = (data: IListItems): IListItems => {
  !data.Select ? (data.Select = "*") : "";
  !data.Topcount ? (data.Topcount = 5000) : "";
  !data.Orderby ? (data.Orderby = "ID") : "";
  !data.Expand ? (data.Expand = "") : "";
  !data.Orderbydecorasc ? (data.Orderbydecorasc = true) : "";
  !data.PageCount ? (data.PageCount = 10) : "";
  !data.PageNumber ? (data.PageNumber = 1) : "";

  return data;
};
const formatFilterValue = (params: IFilter[]): string => {
  let strFilter: string = "";
  for (let i = 0; i < params.length; i++) {
    if (params[i].FilterKey) {
      if (i != 0) {
        strFilter += " and ";
      }

      if (
        params[i].Operator.toLocaleLowerCase() == "eq" ||
        params[i].Operator.toLocaleLowerCase() == "neq" ||
        params[i].Operator.toLocaleLowerCase() == "gt" ||
        params[i].Operator.toLocaleLowerCase() == "lt" ||
        params[i].Operator.toLocaleLowerCase() == "ge" ||
        params[i].Operator.toLocaleLowerCase() == "le"
      )
        strFilter +=
          params[i].FilterKey +
          " " +
          params[i].Operator +
          "'" +
          params[i].FilterValue +
          "'";
      else if (params[i].Operator.toLocaleLowerCase() == "substringof")
        strFilter +=
          params[i].Operator +
          "('" +
          params[i].FilterKey +
          "','" +
          params[i].FilterValue +
          "')";
    }
  }
  return strFilter;
};

export default { SPAddItem, SPUpdateItem, SPDeleteItem, SPReadItems };
