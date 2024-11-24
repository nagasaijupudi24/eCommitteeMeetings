/* eslint-disable @typescript-eslint/no-explicit-any */
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IXenWpCommitteeMeetingsFormsProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  sp:any;
  listName:any;
  committeeMeetingNameList:any;
  formType:string;
  libraryId:any;
  homePageUrl:any;
  passCodeUrl:any
}
