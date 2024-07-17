async function withdrawAllInvitations(maxWithdrawals = Infinity) {
    const withdrawButtons = document.querySelectorAll('button[aria-label^="Withdraw invitation sent to"]');
    let withdrawalsCount = 0;

    for (const button of withdrawButtons) {
        if (withdrawalsCount >= maxWithdrawals) break;

        button.click();
        
        // Wait for the dialog to appear, with some variability
        await new Promise(resolve => setTimeout(resolve, 500 + Math.random() * 1000));

        const confirmButton = [...document.querySelectorAll('.artdeco-modal__confirm-dialog-btn')].find(btn => 
            btn.querySelector('span')?.innerText.includes('Withdraw')
        );
        if (confirmButton) {
            confirmButton.click();
            withdrawalsCount++;
            
            // Wait for the dialog to be processed, with some variability
            await new Promise(resolve => setTimeout(resolve, 500 + Math.random() * 1000));
        }
    }

    console.log(`Withdrawn ${withdrawalsCount} invitations`);
}
