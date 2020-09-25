import { WebPartContext } from "@microsoft/sp-webpart-base";
import  {Web} from "sp-pnp-js";
export interface ITaskActivityProps {
  description: string;
  context: WebPartContext;
  spWeb:Web;
  siteUrl:string;
}
