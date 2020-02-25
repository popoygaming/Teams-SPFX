import { IProfile } from "./IProfile";
import { IMail } from "./IMail";
import { INews } from "./INews";

export interface IHelloWorldState {
    profile: IProfile;
    mailsArray: Array<IMail>;
    newsArray: Array<INews>;
  }
  