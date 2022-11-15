// import { Logger, LogLevel } from "@pnp/logging";
// import { spfi, SPFx, ISPFXContext, SPFI } from "@pnp/sp";
// import { IWeb, Web } from "@pnp/sp/webs";
// import { IRenderListDataParameters } from "@pnp/sp/lists";
// import { ITimeAccounts, TimeAccount } from "../models/ITimeAccount";

// export interface ITimeAccountSPOService {
//     init(context: ISPFXContext): void;
//     getTimeAccounts(): Promise<ITimeAccounts>;
//   }
  
//   export class TimeAccountSPOService implements ITimeAccountSPOService {
//     private LOG_SOURCE: string = "🔶 TimeAccountSPOService";
//     private web: IWeb | undefined;

//     public init(context: ISPFXContext): void {
//         try {
//             // Get the web
//             this.web = spfi().using(SPFx(context)).web;
//         } catch (err) {
//           Logger.write(
//             `${this.LOG_SOURCE} (init) - ${err.message}`,
//             LogLevel.Error
//           );
//         }
//       }

//       public async getTimeAccounts(): Promise<ITimeAccounts> {
//         try {
//             const renderListDataParams: IRenderListDataParameters = {
//                 //ViewXml: "<View><RowLimit>1</RowLimit></View>",
//                 ViewXml: `<View>
//                 <Query>
//                     <Where>
//                         <Eq>
//                             <FieldRef Name='ShowInCard'/>
//                             <Value Type='Boolean'>1</Value>
//                         </Eq>
//                     </Where>
//                 </Query>
//             </View>`
            
//             //Read more: https://www.sharepointdiary.com/2017/04/caml-query-to-filter-yes-no-field-values-in-powershell.html#ixzz7jVDzxphD`, 
//             };
    
    
//             const data = await this.web.lists.getByTitle('TimeOffConfig').renderListDataAsStream(renderListDataParams);
//             const rows = data.Row;
//             let iTimeAccounts: ITimeAccounts = { 
//               timeAccounts: []
//             };
//             for (let index = 0; index < rows.length; index++) 
//             {
//                 const row = rows[index];
//                 if (row['ShowInCard.value'] == '1')
//                 {
//                     const picture = `${row.HolidayTypeIcon.serverUrl}${row.HolidayTypeIcon.serverRelativeUrl}`; 
//                     const timeAccount = new TimeAccount(row.ID, row.Title, row.description, row.sapIdentifier, picture, "");
//                     iTimeAccounts.timeAccounts.push(timeAccount);
//                 }
//             }
//             Logger.write(`${this.LOG_SOURCE} (onAction) - ${iTimeAccounts} - ${iTimeAccounts.timeAccounts}`, LogLevel.Info);
//             return iTimeAccounts;
//         } catch (err) {
//           Logger.write(
//             `${this.LOG_SOURCE} (init) - ${err.message}`,
//             LogLevel.Error
//           );
//         }
//       }
//   }

  

//   export const timeAccountSPOService = new TimeAccountSPOService();