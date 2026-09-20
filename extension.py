from flask_socketio import SocketIO, emit, join_room, leave_room
socketio = SocketIO(async_mode='threading')

documents = {
    '1': '',
    '2': '',
    '3': ''
}

@socketio.on('join_channel')
def on_join(data):
    old_room = data.get('old_channel')
    new_room = data['channel']
    
    # Cleanly switch rooms
    if old_room:
        leave_room(old_room)
    join_room(new_room)
    
    # Send the current text for this channel back to the user
    emit('load_document', {'text': documents.get(new_room, '')})

@socketio.on('text_update')
def handle_text_update(data):
    room = data['channel']
    text = data['text']
    
    # Save text state on the server
    documents[room] = text
    
    # Broadcast to other users in the same room
    emit('receive_update', {'text': text}, room=room, include_self=False)