// tl() for the Vue components: what they show follows the interface language when it changes.
// Every lookup during a render reads `version`, so Vue draws the component again after setLocale.

import { ref } from 'vue';
import { onLocaleChange, tl, trackLocaleReads } from '../i18n';

const version = ref(0);
trackLocaleReads(() => void version.value);
onLocaleChange(() => version.value++);

export { tl };
