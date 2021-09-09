export async function saveToStorageAsync(
  key: string,
  data: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    try {
      Office.context.roamingSettings.set(key, data);
      Office.context.roamingSettings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject();
        }
      });
    } catch (error) {
      reject();
    }
  });
}

export async function getFromStorageAsync(key: string): Promise<string> {
  return new Promise((resolve, reject) => {
    try {
      resolve(Office.context.roamingSettings.get(key));
    } catch (error) {
      reject();
    }
  });
}

export async function setSignatureAsync(signature: string): Promise<void> {
  return new Promise((resolve, reject) => {
    try {
      Office.context.mailbox.item?.body.setSignatureAsync(
        signature,
        { coercionType: Office.CoercionType.Html },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject();
          }
        }
      );
    } catch (error) {
      reject();
    }
  });
}
