from flask import Flask

from extension import socketio

from db.db import init_db

from routes.user_routes import user_bp
from routes.auth_routes import auth_bp
from routes.repository_routes import repository_bp
from routes.document_routes import document_bp
from routes.version_routes import version_bp
from routes.search_routes import search_bp


def create_app():

    app = Flask(__name__)

    app.secret_key = "your_super_secret_and_random_string"

    # Initialize SocketIO
    socketio.init_app(app)

    # Register blueprints
    app.register_blueprint(user_bp)
    app.register_blueprint(auth_bp)
    app.register_blueprint(repository_bp)
    app.register_blueprint(document_bp)
    app.register_blueprint(version_bp)
    app.register_blueprint(search_bp)

    # Initialize database
    with app.app_context():
        init_db()

    return app


app = create_app()


if __name__ == "__main__":
    socketio.run(
        app,
        debug=True
    )