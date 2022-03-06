import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx, SPFI } from "@pnp/sp/presets/all";

class SPServices {
  //   private _context: WebPartContext;
  protected sp: SPFI;
  constructor(context: WebPartContext) {
    //   this._context = context;
    this.sp = spfi().using(SPFx(context));
  }

  public async getLists() {
    const lists: any[] = await this.sp.web.lists
      .filter("Hidden eq false and BaseTemplate eq 100")
      .select("Id,Title")();
    return lists;
  }

  public async getListFields(listId: string) {
    const fields: any[] = await this.sp.web.lists
      .getById(listId)
      .fields.filter(
        "Hidden eq false and ReadOnlyField eq false and Title ne 'Content Type' and Title ne 'Attachments'"
      )
      .select("Title,InternalName")();
    return fields;
  }

  public async getListItems(listId: string, selectedFields: any[]) {
    let expand: any[] = [];
    let select: any[] = ["Id"];
    let items = [];

    for (let i = 0; i < selectedFields.length; i++) {
      switch (selectedFields[i].type) {
        case "SP.FieldUser":
          expand.push(selectedFields[i].key);
          select.push(
            `${selectedFields[i].key}/EMail, ${selectedFields[i].key}/Title,${selectedFields[i].key}/Name`
          );
          break;
        case "SP.FieldLookup":
          expand.push(selectedFields[i].key);
          select.push(`${selectedFields[i].key}/Title`);
          break;
        default:
          select.push(selectedFields[i].key);
          break;
      }
    }


    items = await this.sp.web.lists
      .getById(listId)
      .items.expand(expand.join())
      .select(select.join())();

    return items;
  }
}

export default SPServices;
