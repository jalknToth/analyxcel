from flask import Flask
from controllers.excel_controller import excel_bp
from config import Config

def create_app():

    app = Flask(__name__)
    
    # Load configuration
    app.config.from_object(Config)
    
    # Register blueprints
    app.register_blueprint(excel_bp, url_prefix='/excel')
    
    return app

# Main entry point
if __name__ == '__main__':
    app = create_app()
    app.run(debug=True)