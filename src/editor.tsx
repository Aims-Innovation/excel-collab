import {
  ExcelEditor,
  EditorProps,
  StateContext,
  StateContextValue,
} from './containers';
import {
  memo,
  useEffect,
  useRef,
  useState,
  useSyncExternalStore,
} from 'react';
import { initController } from './controller';
import Worker from './worker?worker&inline';
import { Doc } from 'yjs';
import i18n from './i18n';
import { RTL_LANGUAGE_LIST } from './util/constant';

export type DocConfig = ConstructorParameters<typeof Doc>[0];

export type ExcelProps = EditorProps & {
  doc?: Doc;
  docConfig?: DocConfig;
} & Pick<StateContextValue, 'provider' | 'awareness'>;

const getLanguageSnapshot = () => i18n.current;
const getLanguageServerSnapshot = () => 'en-US';

export const Excel: React.FunctionComponent<ExcelProps> = memo((props) => {
  const { doc, provider, awareness, docConfig } = props;

  const [value, setValue] = useState<StateContextValue | undefined>(undefined);
  const didBootstrapRef = useRef(false);
  const language = useSyncExternalStore(
    i18n.subscribe,
    getLanguageSnapshot,
    getLanguageServerSnapshot,
  );

  useEffect(() => {
    if (didBootstrapRef.current) {
      return;
    }
    didBootstrapRef.current = true;
    i18n.init();

    const controller = initController({
      worker: new Worker(),
      doc: doc ? doc : docConfig ? new Doc(docConfig) : new Doc(),
    });

    setValue({ controller, provider, awareness });
  }, []);

  useEffect(() => {
    const isRTL = RTL_LANGUAGE_LIST.includes(language as any);
    document.documentElement.setAttribute(
      'data-layout-direction',
      isRTL ? 'rtl' : 'ltr',
    );
  }, [language]);

  return (
    <StateContext.Provider value={value}>
      {value && <ExcelEditor {...props} />}
    </StateContext.Provider>
  );
});

Excel.displayName = 'Excel';
