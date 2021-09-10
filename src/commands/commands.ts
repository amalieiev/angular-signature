import {
  ACCEESS_TOKEN_STORAGE_KEY,
  SIGNATURE_STORAGE_KEY,
} from '../app/shared/constants';
import {
  getAccessTokenAsync,
  getFromStorageAsync,
  saveToStorageAsync,
  setSignatureAsync,
} from '../app/shared/office';

export async function onMessageComposeHandler(event: any): Promise<void> {
  // await insertSignatureOnCompose();
  // await requestAccessTokenAsync();
  event.completed();
}

export async function insertSignatureOnCompose() {
  const signature = await getFromStorageAsync(SIGNATURE_STORAGE_KEY);
  if (signature) {
    await setSignatureAsync(signature);
  }
}

export async function requestAccessTokenAsync(): Promise<void> {
  const token = await getAccessTokenAsync();
  console.log('token', token);
  await saveToStorageAsync(ACCEESS_TOKEN_STORAGE_KEY, token);
}

function getGlobal() {
  if (typeof self !== 'undefined') {
    return self;
  }
  if (typeof window !== 'undefined') {
    return window;
  }
  return typeof global !== 'undefined' ? global : undefined;
}

const g = getGlobal() as any;

g.onMessageComposeHandler = onMessageComposeHandler;

Office.actions.associate('onMessageComposeHandler', onMessageComposeHandler);
