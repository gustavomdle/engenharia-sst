import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISstNovoProjectProps {
  description: string;
  context: WebPartContext;
  siteurl: string;
}
