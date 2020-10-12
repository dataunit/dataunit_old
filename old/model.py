class TestCase:

    def __init__(self, test_case_id, active=True, name=None, description=None):
        self.test_case_id = test_case_id
        self.active = active
        self.name = name
        self.description = description

