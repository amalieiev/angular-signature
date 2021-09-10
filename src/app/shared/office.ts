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

export const getAccessTokenAsync = (
  options = { allowSignInPrompt: true, allowConsentPrompt: true }
): Promise<string> =>
  new Promise<string>((resolve, reject) => {
    const ACCESS_TOKEN_TIMEOUT = 1000 * 60 * 5; // 5 minutes.
    const errorMessage =
      'We were unable to sign you into Microsoft. Click below to try again.';
    const timeout = setTimeout(() => {
      reject(errorMessage);
    }, ACCESS_TOKEN_TIMEOUT);

    (async (): Promise<void> => {
      try {
        const accessToken = await OfficeRuntime.auth.getAccessToken(options);
        clearTimeout(timeout);
        resolve(accessToken);
      } catch (error: any) {
        clearTimeout(timeout);
        reject(error.message);
      }
    })();
  });
