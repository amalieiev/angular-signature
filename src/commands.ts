import { SIGNATURE_STORAGE_KEY } from './constants';
import { getFromStorageAsync, setSignatureAsync } from './office';

export async function onMessageComposeHandler(event: any): Promise<void> {
  await insertSignatureOnCompose();
  event.completed();
}

export async function insertSignatureOnCompose() {
  const signature = await getFromStorageAsync(SIGNATURE_STORAGE_KEY);
  if (signature) {
    await setSignatureAsync(signature);
  }
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
