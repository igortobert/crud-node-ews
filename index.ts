process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';
import EWS from 'node-ews';

export class Calendar {
  protected HandlerEWS: EWS
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
      if (result.ResponseMessages.FindItemResponseMessage.RootFolder.Items) {
        let calendarItems: any = result.ResponseMessages.FindItemResponseMessage.RootFolder.Items.CalendarItem
        if (!Array.isArray(calendarItems)) calendarItems = [calendarItems]
        calendarItems.length ? calendarItems = calendarItems : calendarItems = []
        return calendarItems
      } else {
        return []
      }

    } catch (error) {
      return error
    }
  }

  async create(meeting: Meeting, attendee?: string[]): Promise<Attributes> {
    const ewsArgs = {
      attributes: {
        SendMeetingInvitations: attendee && attendee.length ? 'SendToAllAndSaveCopy' : 'SendToNone'
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
          ReminderMinutesBeforeStart: meeting.ReminderMinutes ? meeting.ReminderMinutes.toString() : '0',
          Start: meeting.Start,
          End: meeting.End,
          Location: meeting.Location,
          RequiredAttendees: {
            Attendee: attendee && attendee.length ? attendee.map(v => ({ Mailbox: { EmailAddress: v } })) : []
          }
        }
      }
    }

    try {
      const resp = await this.HandlerEWS.run('CreateItem', ewsArgs)
      return resp.ResponseMessages.CreateItemResponseMessage.Items.CalendarItem.ItemId.attributes
    } catch (err) {
      return err.message;
    }
  }
  // async create(meeting: Meeting): Promise<Attributes> {
  //   const ewsArgs = {
  //     attributes: {
  //       SendMeetingInvitations: 'SendToNone'
  //     },
  //     Items: {
  //       CalendarItem: {
  //         Subject: meeting.Subject,
  //         Body: {
  //           attributes: {
  //             BodyType: "HTML"
  //           },
  //           "$value": meeting.Body
  //         },
  //         ReminderMinutesBeforeStart: meeting.ReminderMinutes ? meeting.ReminderMinutes.toString() : '0',
  //         Start: meeting.Start,
  //         End: meeting.End,
  //         Location: meeting.Location
  //       }
  //     }
  //   }

  //   try {
  //     const resp = await this.HandlerEWS.run('CreateItem', ewsArgs)
  //     return resp.ResponseMessages.CreateItemResponseMessage.Items.CalendarItem.ItemId.attributes
  //   } catch (err) {
  //     return err.message;
  //   }
  // }

  async update(attributes: Attributes, meeting: Meeting) {
    let ewsArgs: any = {
      attributes: {
        MessageDisposition: 'SaveOnly',
        ConflictResolution: 'AlwaysOverwrite',
        SendMeetingInvitationsOrCancellations: 'SendToAllAndSaveCopy'
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
    let agrsArray: SetItemField[] = []
    for (let k in meeting) {
      let obj: SetItemField = null
      if (meeting[k].length) {
        switch (k) {
          case 'Subject':
            obj = {
              FieldURI: { attributes: { FieldURI: `item:${k}` } },
              CalendarItem: { [k]: meeting[k] }
            }
            break;
          case 'Start':
          case 'End':
          case 'Location':
            obj = {
              FieldURI: { attributes: { FieldURI: `calendar:${k}` } },
              CalendarItem: { [k]: meeting[k] }
            }
            break;
        }
        if (obj && Object.keys(obj).length) agrsArray.push(obj)
      }
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
        DeleteType: 'MoveToDeletedItems',
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

export interface Meeting {
  Subject?: string
  Body?: string
  ReminderMinutes?: number
  Start?: string
  End?: string
  Location?: string
}

export interface Attributes {
  Id: string
  ChangeKey: string
}

interface SetItemField {
  FieldURI: {
    attributes: {
      FieldURI: string
    }
  }
  CalendarItem: Meeting
}









