import { isArray } from "util";

process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';
import EWS from 'node-ews';

const URL = "https://192.168.87.29"
const credentials = {
  userName: "test_user",
  password: "Qwerty123456789"
}



export class Calendar {
  protected HandlerEWS
  constructor(
    private readonly HOST: string,
    private readonly USER: string,
    private readonly PASS: string,
  ) {
    this.HandlerEWS = new EWS({
      host: HOST,
      username: USER,
      password: PASS,
    }, {});
  }

  private async getCalendarAttributes() {
    let ewsArgs: any = {
      FolderShape: {
        BaseShape: 'Default'
      },
      FolderIds: {
        DistinguishedFolderId: {
          attributes: {
            Id: 'calendar'
          }
        }
      }
    }

    try {
      let result = await this.HandlerEWS.run('GetFolder', ewsArgs)
      let attributes = result.ResponseMessages.GetFolderResponseMessage.Folders.CalendarFolder.FolderId.attributes
      return attributes
    } catch (err) {
      return err
    }
  }

  async findAll(MaxEntries: number, StartDate: string, EndDate: string) {
    const attributes: any = await this.getCalendarAttributes()
    const ewsArgs = {
      attributes: {
        Traversal: 'Shallow'
      },
      ItemShape: {
        BaseShape: 'Default'
      },
      CalendarView: {
        attributes: {
          MaxEntriesReturned: MaxEntries,
          StartDate,
          EndDate
        }
      },
      ParentFolderIds: {
        FolderId: {
          attributes
        }
      }
    }

    try {

      const result = await this.HandlerEWS.run('FindItem', ewsArgs)
      let calendarItems: any = result.ResponseMessages.FindItemResponseMessage.RootFolder.Items.CalendarItem
      if (!isArray(calendarItems)) calendarItems = [calendarItems]
      calendarItems.length ? calendarItems = calendarItems : calendarItems = []
      return calendarItems

    } catch (error) {
      return error
    }
  }

  async create(meeting: CreateMeeting): Promise<Attributes> {
    const ewsArgs = {
      attributes: {
        SendMeetingInvitations: 'SendToNone'
      },
      Items: {
        CalendarItem: {
          Subject: meeting.Subject,
          Body: {
            attributes: {
              BodyType: "HTML"
            },
            "$value": meeting.Body
          },
          ReminderMinutesBeforeStart: meeting.ReminderMinutes.toString(),
          Start: meeting.Start,
          End: meeting.End,
          Location: meeting.Location
        }
      }
    }

    try {
      const resp = await this.HandlerEWS.run('CreateItem', ewsArgs)
      return resp.ResponseMessages.CreateItemResponseMessage.Items.CalendarItem.ItemId
    } catch (err) {
      return err.message;
    }
  }

  async update(attributes: Attributes, meeting: CreateMeeting) {
    let ewsArgs: any = {
      attributes: {
        MessageDisposition: 'SaveOnly',
        ConflictResolution: 'AlwaysOverwrite',
        SendMeetingInvitationsOrCancellations: 'SendToNone'
      },
      ItemChanges: {
        ItemChange: {
          ItemId: {
            attributes
          },
          Updates: {

          }
        }
      }
    }
    let agrsArray: any = []
    if (meeting.Subject) {
      let subject = {
        FieldURI: { attributes: { FieldURI: 'item:Subject' } },
        CalendarItem: { Subject: meeting.Subject }
      }
      agrsArray.push(subject)
    }

    if (meeting.Location) {
      let location = {
        FieldURI: { attributes: { FieldURI: 'calendar:Location' } },
        CalendarItem: { Location: meeting.Location }
      }
      agrsArray.push(location)
    }

    ewsArgs.ItemChanges.ItemChange.Updates.SetItemField = agrsArray

    try {
      const resp = await this.HandlerEWS.run('UpdateItem', ewsArgs)
      return resp.ResponseMessages.UpdateItemResponseMessage
    } catch (err) {
      return err.message;
    }

  }

  async delete(attributes: Attributes): Promise<string> {
    const ewsArgs = {
      attributes: {
        DeleteType: 'HardDelete',
        SendMeetingCancellations: 'SendToAllAndSaveCopy'
      },
      ItemIds: {
        ItemId: {
          attributes
        }
      }
    }

    try {
      const resp = await this.HandlerEWS.run('DeleteItem', ewsArgs)
      return resp.ResponseMessages.DeleteItemResponseMessage
    } catch (err) {
      return err.message;
    }
  }
}

export interface CreateMeeting {
  Subject: string
  Body: string
  ReminderMinutes: number
  Start: string
  End: string
  Location: string
}

export interface Attributes {
  Id: string
  ChangeKey: string
}











