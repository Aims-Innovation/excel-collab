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
import { reactLog } from './util/debug';
import {
  sizeConfig,
  lightColor,
  darkColor,
  getTheme,
} from './theme';

export type DocConfig = ConstructorParameters<typeof Doc>[0];

export type ExcelProps = EditorProps & {
  doc?: Doc;
  docConfig?: DocConfig;
} & Pick<StateContextValue, 'provider' | 'awareness'>;

const getLanguageSnapshot = () => i18n.current;
const getLanguageServerSnapshot = () => 'en-US';

/**
 * Push our theme tokens (colors + sizes + zIndex) onto the document
 * root as CSS custom properties. Previously this lived inside the
 * MenuBar's <Theme> widget, which meant consumers passing
 * `hideTheme={true}` (introduced in v0.1.13.17 to drop the toggle
 * button) silently lost every `var(--borderColor)` / `var(--white)` /
 * etc. reference -- borders, hover backgrounds, and the formula bar
 * separator all stopped rendering. The canvas grid kept working
 * because it draws via the canvas API, not CSS, which is what
 * masked the bug for so long. Bootstrapping the variables here
 * keeps the toggle widget's runtime light/dark switching intact
 * while guaranteeing the variables exist whether or not the widget
 * is mounted.
 */
function bootstrapCssVariables() {
  if (typeof document === 'undefined') return;
  const root = document.documentElement;
  const set = (data: Record<string, string | number>) => {
    for (const key of Object.keys(data)) {
      root.style.setProperty(`--${key}`, String(data[key] ?? ''));
    }
  };
  set(sizeConfig);
  set(getTheme() === 'dark' ? darkColor : lightColor);
}

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
    reactLog('Excel bootstrap: creating Worker + controller');
    i18n.init();
    bootstrapCssVariables();

    const controller = initController({
      worker: new Worker(),
      doc: doc ? doc : docConfig ? new Doc(docConfig) : new Doc(),
    });

    setValue({ controller, provider, awareness });
  }, []);

  useEffect(() => {
    reactLog('language changed', language);
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
