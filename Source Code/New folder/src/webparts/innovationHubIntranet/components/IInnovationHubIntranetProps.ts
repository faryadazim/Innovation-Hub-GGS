import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IInnovationHubIntranetProps {
  context: WebPartContext;
}

export interface IPeoplelist{
  key: number;
  imageUrl:string;
  text: string;
  ID: number;
  secondaryText: string;
  isValid: true;
}

export interface IDropdownOption{
  key:string;
  text:string;
}