import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IRenderListProps {
  context: WebPartContext;
  list: string;
  fields: any[];
  title: string;
}
