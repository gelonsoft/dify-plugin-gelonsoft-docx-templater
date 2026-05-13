from dify_plugin import ToolProvider


class DocxTemplaterProvider(ToolProvider):
    def _validate_credentials(self, credentials: dict) -> None:
        """
        Validate provider credentials.
        This plugin does not require any credentials.
        """
        pass
