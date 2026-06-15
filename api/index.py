from app.main import DemoHandler, init_db, load_env_file


load_env_file()
init_db()


class handler(DemoHandler):
    pass
