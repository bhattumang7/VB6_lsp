
# Git 
Do not mention any AI model in the commits. 
Do not push untill you get express approval from me. Do not switch the branches without approval from me.

# Contribur
Contributor name is "Umang Bhatt" and controbutor email is bhatt.umang7@gmail.com in the commits.
	
# naming the bugs 
When the user reports bugs, they might share them with numering (either 1,2,3 or A B C etc). Do not mention Bug 1 or Bug 2 or anything similar in places. Those numbers have no meaning when we look back those are just for the conversation.

# Known gaps 
Known gaps are not acceptable. We must do it right. Known divergence are not okay. Do not put any dummy/assumed value anywhere which we have to come back annd clear later. If neded created methods with notimplemented exceptions. Never mark a test as #[ignore] if something is not behaving correctly - this is a shortcut and dont take shortcuts.

# Test design
The tests must verify the exact ast and the exact parts of AST and not just presence of the error. Verify all 	 parts of the AST
Keep the tests in a separate folder, not in the same file as of the System Under Test (SUT).

# test coverage 
Make sure that the tests cover all the changes we are making. 

# private sources
Do not mention anything in the comments that could lead to thinking "where did they figure this out from?" - all the references need to remain private. i.e.  Direct port of `EbProcessLinkedList` or similar things are not acceptable. the material referenced (whether the code or docs or experiments) needs to remain private and should not be revealed about in this repo.

