
# Git 
Do not mention any AI model in the commits. 

# Contribur
Contributor name is "Umang Bhatt" and controbutor email is bhatt.umang7@gmail.com in the commits.

# naming the bugs 
When the user reports bugs, they might share them with numering (either 1,2,3 or A B C etc). Do not mention Bug 1 or Bug 2 or anything similar in places. Those numbers have no meaning when we look back those are just for the conversation.

# Known gaps 
Known gaps are not acceptable. We must do it right. Known divergence are not okay. Do not put any dummy/assumed value anywhere which we have to come back annd clear later. If neded created methods with notimplemented exceptions. Never mark a test as #[ignore] if something is not behaving correctly - this is a shortcut and dont take shortcuts.

# Test design
The tests must verify the exact ast and the exact parts of AST and not just presence of the error. Verify all the parts of the AST.