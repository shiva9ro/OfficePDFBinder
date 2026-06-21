import os
from pathlib import Path

from PySide6.QtCore import QCoreApplication, QLocale, QTranslator


DEFAULT_LANGUAGE = "ja"
SUPPORTED_LANGUAGES = ("ja", "en")
LANGUAGE_FILENAME = "OfficePDFBinder.language"
TRANSLATION_BASENAME = "OfficePDFBinder"

_active_translator = None
_active_language = DEFAULT_LANGUAGE


def normalize_language(value):
    """言語名、ロケール名、Inno Setupの言語名をja/enへ正規化する。"""
    normalized = (value or "").strip().lower().replace("-", "_")
    if normalized in {"ja", "ja_jp", "japanese"}:
        return "ja"
    if normalized in {"en", "en_us", "en_gb", "english"}:
        return "en"
    if normalized.startswith("ja_"):
        return "ja"
    if normalized.startswith("en_"):
        return "en"
    return None


def resolve_language(runtime_dir, environ=None, system_locale=None):
    """環境変数、インストーラー設定、OS表示言語の順に使用言語を決める。"""
    environ = os.environ if environ is None else environ
    language = normalize_language(environ.get("OFFICEPDFBINDER_LANGUAGE"))
    if language:
        return language

    language_path = Path(runtime_dir) / LANGUAGE_FILENAME
    try:
        language = normalize_language(language_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError):
        language = None
    if language:
        return language

    locale_name = QLocale.system().name() if system_locale is None else system_locale
    language = normalize_language(locale_name)
    if language:
        return language
    # 対応言語が日本語と英語だけの間は、日本語以外のOSでは英語を使用する。
    return "en"


def install_app_translator(app, runtime_dir):
    """選択された言語のQt翻訳を読み込み、言語コードと成否を返す。"""
    global _active_language, _active_translator

    language = resolve_language(runtime_dir)
    _active_translator = None
    _active_language = DEFAULT_LANGUAGE
    if language == DEFAULT_LANGUAGE:
        return language, True

    translation_path = (
        Path(runtime_dir)
        / "translations"
        / f"{TRANSLATION_BASENAME}_{language}.qm"
    )
    translator = QTranslator(app)
    if not translator.load(str(translation_path)):
        return DEFAULT_LANGUAGE, False

    app.installTranslator(translator)
    _active_translator = translator
    _active_language = language
    return language, True


def current_language():
    return _active_language


def translate(context, source_text):
    """QObjectを継承しない処理からも利用できる共通翻訳関数。"""
    return QCoreApplication.translate(context, source_text)
