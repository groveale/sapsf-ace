import { ITimeBooked } from "./ITimeBooked";

export interface ITimeAccount {
    id: number;
    title: string;
    description: string;
    sapIdentifierTAT: string;
    sapIdentifierTT: string;
    picture: string;
    base64: string;
    balanceDays: number;
    blanaceHours: number;
    balanceDaysString: string;
    balanceHoursString: string;
    timeBookedPast: ITimeBooked[];
    timeBookedUpcoming: ITimeBooked[];
  }