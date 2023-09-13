from flask import Flask, request, jsonify
import os

app = Flask(__name__)

# Define a route to handle POST requests with multipart form data
@app.route('/zoho/file/uploader', methods=['POST'])
def handle_upload():
    try:
        # Check if the 'file' field is in the request
        if 'file' not in request.files:
            return jsonify({"error": "No file part"}), 400
        
        file = request.files['file']
        
        # Check if the file is empty
        if file.filename == '':
            return jsonify({"error": "No selected file"}), 400
        else:
            print("File read from request = " + file.filename)
        
        # Save the uploaded file to a designated directory
        root_dir = os.path.dirname(os.path.abspath(__file__))

        dest_file_path = os.path.join(root_dir, file.filename)

        file.save(dest_file_path)

        print("File saved in file path = " + dest_file_path)
        
        # You can also access other form fields if needed. 
        # For example assume filename and format key sent along with 
        file_name = request.form.get('filename')
        format = request.form.get('format')

        print("file_name read from request body = " + str(file_name))
        print("format from from request body = " + str(format))

        # Process the uploaded file and form data as needed

        # Send a response back to Zoho API to acknowledge the callback successfully completed.
        # String value given for message key below will shown as save status in editor UI.
        response_data = {"message": "File uploaded successfully"}

        return jsonify(response_data), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 400

print('\nMake a post request to this end point to test save callback from local file - http://localhost/${port}${callbackEndpoint}')
print('\nKeep the multipart key as content for file. If you send the file in different key then change the key in save_callback code as well before testing')
print('\nYou should not use http://localhost/${port} as save_url in callback settings. You should only use it as publicly accessible(with your application domain) end point.')

if __name__ == '__main__':
    app.run(debug=True)
