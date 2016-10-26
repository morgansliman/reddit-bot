import praw

user_agent = 'karma checker v1'
r = praw.Reddit(user_agent = user_agent)

# Returns redditor(username)'s total karma in given subreddit
def getKarmaBySub(username, subreddit):
    user = r.get_redditor(username)
    
    gen = user.get_submitted(limit=None)
    karma = {subreddit : 0}
    for thing in gen:
        if thing.subreddit.display_name == subreddit:
            karma[subreddit] = (karma.get(subreddit) + thing.score)
    return karma[subreddit]
